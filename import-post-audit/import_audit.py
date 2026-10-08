#!/usr/bin/env python3
"""輸入事後調査資料の作成ツール。

輸入許可通知書(PDF)から申告番号・輸入許可年月日・仕出人を読み取り、
  1. 管理番号(No.)を採番
  2. 輸入明細一覧(Excel)へ行を追加
  3. 「表紙」(Excel)を作成
してデスクトップの作業フォルダに保存する。

使い方:
  python import_audit.py 許可通知書.pdf [...] --office "林六／東京"            # 海外送金あり
  python import_audit.py 許可通知書.pdf --nosend 着払 --office "林六／東京"     # 海外送金なし(着払/無償/乙仲)
  python import_audit.py フォルダ名 --office ...                                # フォルダ内の PDF すべて

送金日・金額・乙仲業者名など PDF にない項目は、台帳の空欄に手入力する。
"""
import argparse
import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import pdfplumber
from openpyxl.styles import PatternFill

HERE = Path(__file__).resolve().parent
TPL_LEDGER = HERE / "templates" / "輸入明細一覧テンプレート.xlsx"
TPL_COVER = HERE / "templates" / "表紙テンプレート.xlsx"
VERSION = "2026-10-08 l (表紙のひな形をシート名で探す)"
SHEET_SEND = "海外送金あり"
SHEET_NOSEND = "海外送金なし（乙仲・無償・着払）"
HILITE = PatternFill("solid", fgColor="FFFF00")

# 金額(送金額・仕入金額・単価・表紙の金額)と、送金計算書・海外送金依頼書の読み取りは使わない。
# True にすると、送金書類から金額まで読み取って結びつける機能が有効になる。
READ_AMOUNTS = False

# 同じ月(輸入許可日の年月)に複数あっても、管理番号を1つにまとめる仕出人
MERGE_MONTHLY = ["TAEKWANG INDUSTRIAL CO.,LTD."]


def open_template(path: Path, embedded: str):
    """templates フォルダのひな形があればそれを、無ければ埋め込みのひな形を使う。"""
    import base64
    import io
    if path.exists():
        return path
    return io.BytesIO(base64.b64decode("".join(embedded.split())))


# ---------- 期・出力先 ----------
def fiscal_period(d: datetime) -> int:
    """4月始まり。2026/4〜2027/3 が第81期。"""
    y = d.year if d.month >= 4 else d.year - 1
    return y - 1945


def period_label(n: int) -> str:
    y = n + 1945
    return f"第{n}期（{y}.4月～{y + 1}.3月）"


def desktop() -> Path:
    home = Path.home()
    for p in (home / "Desktop", home / "OneDrive" / "Desktop", home / "デスクトップ"):
        if p.is_dir():
            return p
    return home


# ---------- PDF 読み取り ----------
_ocr_engine = None
DATE_RE = re.compile(r"(20\d{2})[/年.-](\d{1,2})[/月.-](\d{1,2})")
SPACED_DECL = re.compile(r"(?<!\d)(\d{3}) (\d{4}) (\d{4})(?!\d)")
PLAIN_DECL = re.compile(r"(?<!\d)(\d{3})(\d{4})(\d{4})(?!\d)")
COMPANY = re.compile(r"[A-Za-z][A-Za-z0-9 .,&()'/-]{3,}(?:CO|LTD|LIMITED|INC|CORP|GMBH|LLC|PTE|S\.A)[A-Za-z0-9 .,&()'/-]*")
OCR_MAX_PAGES = 10


def ocr_items(pdf: Path, page_no: int):
    """スキャンPDFの指定ページを OCR し、[(y比率, x, text)] を返す。"""
    global _ocr_engine
    try:
        import pypdfium2 as pdfium
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        raise ValueError("スキャンPDFの読み取りには `pip install rapidocr-onnxruntime pypdfium2` が必要です")
    if _ocr_engine is None:
        _ocr_engine = RapidOCR()
    img = pdfium.PdfDocument(str(pdf))[page_no].render(scale=3).to_numpy()
    res, _ = _ocr_engine(img)
    h = img.shape[0]
    return [(b[0][1] / h, b[0][0], t) for b, t, _ in (res or [])]


def norm_name(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", s.upper().replace("LID", "LTD"))


def match_supplier(raw: str, suppliers) -> tuple[str, bool]:
    """OCRの崩れた仕出人名を既知の仕入先名に寄せる。(名前, 確実か) を返す。"""
    import difflib
    key = norm_name(raw)
    best, score = raw, 0.0
    for name in suppliers:
        r = difflib.SequenceMatcher(None, key, norm_name(name)).ratio()
        if r > score:
            best, score = name, r
    return (best, True) if score >= 0.85 else (raw, False)


def find_known_supplier(text: str, suppliers) -> str | None:
    """ページ全体に既知の仕入先名が(表記ゆれ込みで)含まれていればその名前。"""
    flat = norm_name(text)
    hits = [n for n in suppliers if len(norm_name(n)) >= 8 and norm_name(n) in flat]
    return max(hits, key=len) if hits else None


def pick_date(text: str, labels=("輸入許可日", "審査終了日")):
    for lab in labels:
        m = re.search(lab + r"[\s\\|:：]*" + DATE_RE.pattern, text)
        if m:
            return datetime(*map(int, m.groups()))
    return None


def parse_page_text(text: str, suppliers) -> dict | None:
    """文字情報のあるページ1枚分。許可通知書でなければ None。"""
    if not re.search(r"輸入許可|SEA/IMP|AIR/IMP", text):
        return None
    flat = re.sub(r"[ \u3000]+", " ", text)
    m = SPACED_DECL.search(flat)
    if not m:  # 「輸入申告番号等」の11桁
        m = re.search(r"輸入申告番号等\s*" + PLAIN_DECL.pattern, flat)
    decl = " ".join(m.groups()) if m else None
    if not decl:
        return None
    permit = pick_date(flat)
    if not permit and (d := DATE_RE.search(flat)):
        permit = datetime(*map(int, d.groups()))
    shipper = ""
    m = re.search(r"仕\s*出\s*人[^A-Za-z]{0,40}(" + COMPANY.pattern + ")", flat)
    if m:
        shipper = m.group(1).strip(" -.,")
    return {"decl": decl, "permit": permit, "shipper": shipper}


def parse_items(items, suppliers) -> dict | None:
    """OCR結果(1ページ)。ラベルが誤認識されても数字・位置で拾う。"""
    items = sorted(items)
    head = [t for y, _, t in items if y < 0.2]
    decl = next((" ".join(m.groups()) for t in head if (m := SPACED_DECL.search(t))), None)
    if not decl:
        return None  # 許可通知書の見出しがないページ(納付通知など)は対象外
    permit = None
    for lo in (0.8, 0.0):  # 下部の輸入許可日 → なければ上部の申告年月日
        for y, _, t in items:
            if y >= lo and (m := DATE_RE.search(t)):
                permit = datetime(*map(int, m.groups()))
                break
        if permit:
            break
    shipper = ""
    for y, x, t in items:  # 「出人」ラベルの右隣
        if re.search(r"仕\s*出\s*人|出\s*人$", t) and len(t) <= 6:
            cand = [(xx, tt) for yy, xx, tt in items if abs(yy - y) < 0.012 and xx > x + 50 and re.search(r"[A-Za-z]{3}", tt)]
            if cand:
                shipper = min(cand)[1]
                break
    return {"decl": decl, "permit": permit, "shipper": shipper}


def read_permits(pdf: Path, suppliers=()) -> list[dict]:
    """1つのPDFから許可通知書を全ページ分読む(複数申告・納付通知付きにも対応)。"""
    found: dict[str, dict] = {}
    with pdfplumber.open(pdf) as doc:
        texts = [(p.extract_text(layout=True) or "") for p in doc.pages]
    for i, t in enumerate(texts):
        info = parse_page_text(t, suppliers) if t.strip() else None
        if info is None and not t.strip() and i < OCR_MAX_PAGES:  # 文字なしページは OCR
            items = ocr_items(pdf, i)
            info = parse_items(items, suppliers)
            t = "\n".join(x[2] for x in items)
        if info is None:
            continue
        if not info["shipper"] or not find_known_supplier(info["shipper"], suppliers):
            known = find_known_supplier(t, suppliers)
            if known:
                info["shipper"] = known
        found.setdefault(info["decl"], info)
    out = []
    for info in found.values():
        sure = True
        if info["shipper"] and suppliers:
            info["shipper"], sure = match_supplier(info["shipper"], suppliers)
        elif not info["shipper"]:
            sure = False
        info["shipper_sure"] = sure
        out.append(info)
    if not out:
        raise ValueError("許可通知書(申告番号)が見つかりません。別の書類か、読み取れない画像の可能性があります")
    return out


# ---------- 追加: 送金計算書・許可通知書の金額/数量、送金との結びつけ ----------
DEFAULT_PRODUCTS = {"HYDROGEN PEROXIDE": "過酸化水素", "CA-801H": "CA801H"}
AMT = r"(\d{1,3}(?:\s*,\s*\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)"


def to_num(s: str) -> float:
    return float(re.sub(r"[\s,]", "", s))


def load_products(root: Path) -> dict[str, str]:
    """商品名の変換表(英語の品名 → 台帳に書く名前)。商品名.txt に「ENGLISH=日本語」で追記できる。"""
    f = root / "商品名.txt"
    if not f.exists():
        root.mkdir(parents=True, exist_ok=True)
        f.write_text("\n".join(f"{k}={v}" for k, v in DEFAULT_PRODUCTS.items()) + "\n", encoding="utf8")
    d = {}
    for line in f.read_text(encoding="utf8").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            if k.strip():
                d[k.strip().upper()] = v.strip()
    return d


def loose_amount(tok: str, cur: str = "JPY") -> float | None:
    """OCRで , が . や : に化けた金額も読む。外貨は末尾2桁を小数、円は整数として扱う。"""
    t = tok.strip()
    m = re.fullmatch(r"([\d.,:; ]*\d)\s*[.,:;]\s*(\d{2})", t)
    if m and (cur != "JPY" or m.group(2) == "00"):
        ip, dec = re.sub(r"\D", "", m.group(1)), m.group(2)
    else:
        ip, dec = re.sub(r"\D", "", t), "00"
    return float(f"{ip}.{dec}") if ip else None


def extras_from_items(items, products: dict, have_amt: bool, have_qty: bool) -> dict:
    """OCRの文字位置から、仕入書価格・数量を拾う(文字列の並びが崩れたときの補完)。"""
    out = {}
    num = re.compile(r"[\d][\d.,:; ]*")
    if not have_amt:
        anchors = [(y, x, t) for y, x, t in items if re.search(r"CIF|FOB|CFR|CIP|C&F", t)]
        for y0, x0, t0 in anchors:
            cur = next((c for c in ("JPY", "USD", "EUR", "CNY", "GBP", "THB", "AUD")
                        if any(c in tt for yy, xx, tt in items if abs(yy - y0) < 0.04 and abs(xx - x0) < 500)), None)
            cands = [(abs(yy - y0) * 3 + abs(xx - x0) / 3000, tt) for yy, xx, tt in items
                     if -0.012 <= yy - y0 <= 0.05 and xx >= x0 - 80 and num.fullmatch(tt.strip()) and len(re.sub(r"\D", "", tt)) >= 3]
            if cur and cands:
                v = loose_amount(min(cands)[1], cur)
                if v:
                    out["inv_cur"], out["inv_amt"] = cur, v
                    break
    if not have_qty:
        cands = []
        for y, x, t in items:
            m = re.fullmatch(r"\s*(\d[\d.,:; ]*)\s*K\s*G\s*", t)
            if m and y > 0.4:
                cands.append((y, loose_amount(m.group(1), "USD")))
        cands = [(y, v) for y, v in cands if v]
        if cands:
            out["qty"], out["unit"] = sorted(cands)[0][1], "Ｋｇ"
    return out


def extract_extras(text: str, products: dict) -> dict:
    """許可通知書から仕入書価格(通貨・金額)・数量・商品名を拾う。"""
    flat = re.sub(r"[ 　]+", " ", text)
    out = {"inv_cur": None, "inv_amt": None, "qty": None, "unit": "", "product": ""}
    m = re.search(r"(?:CIF|FOB|CFR|C&F|CIP)[\s\-–\\]*(JPY|USD|EUR|CNY|GBP|THB|AUD)[\s\-–\\]*" + AMT, flat)
    if m:
        out["inv_cur"], out["inv_amt"] = m.group(1), to_num(m.group(2))
    m = re.search(r"数量\s*(?:[\(（]\s*[12１２]\s*[\)）])?\s*" + AMT + r"\s*KG\b", flat)
    if m and to_num(m.group(1)) > 0:
        out["qty"], out["unit"] = to_num(m.group(1)), "Ｋｇ"
    up = flat.upper()
    for k, v in products.items():
        if k in up:
            out["product"] = v
            break
    return out


def is_remit_text(t: str, name: str = "") -> bool:
    if re.search(r"輸入許可通知|SEA/IMP|AIR/IMP", t):
        return False
    return "送金計算書" in name or bool(re.search(
        r"STATEMENT|外国.{0,2}係?計算|外国送金\(電信\)|REMITTANCE WITH DECLARATION|BANKING CORPORATION", t))


def fix_year(y: int) -> int:
    return y - 800 if 2800 <= y < 2900 else y  # OCRで 0 が 8 に化けた年を直す


def parse_remit(t: str, suppliers) -> dict | None:
    """外国送金の計算書(銀行発行)。送金日・通貨・外貨額・円貨額・受取人を拾う。"""
    from collections import Counter
    import difflib
    flat = re.sub(r"[ 　]+", " ", t)
    date = None
    for pat in (r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", r"(?:DATE|取組日)\D{0,12}(\d{4})/(\d{1,2})/(\d{1,2})"):
        m = re.search(pat, flat)
        if m and 2000 <= fix_year(int(m.group(1))) < 2100:
            date = datetime(fix_year(int(m.group(1))), int(m.group(2)), int(m.group(3)))
            break
    cur = amount = rate = yen = None
    m = re.search(r"(US\$|USD|EUR|GBP|CNY|THB|AUD)\s*(\d{1,3}(?:,\d{3})*\.\d{2})", flat)
    yen_cands = [int(x.replace(",", "")) for x in re.findall(r"[¥\\]\s*(\d{1,3}(?:,\d{3}){2,})", flat)]
    if m:
        cur = "USD" if m.group(1) in ("US$", "USD") else m.group(1)
        amount = to_num(m.group(2))
        r = re.search(r"@\s*(\d{2,3}\.\d{1,4})", flat)
        rate = float(r.group(1)) if r else None
        if rate:
            yen = round(amount * rate)
        elif yen_cands:
            yen = min(yen_cands)
    elif yen_cands:
        cur = "JPY"
        amount = yen = Counter(yen_cands).most_common(1)[0][0]
    if date is None or amount is None:
        return None
    supplier, sure = "", False
    best = 0.0
    for line in flat.splitlines():
        if len(re.findall(r"[A-Za-z]", line)) < 8:
            continue
        for name in suppliers:
            r2 = difflib.SequenceMatcher(None, norm_name(line), norm_name(name)).ratio()
            if r2 > best and r2 >= 0.78:
                best, supplier, sure = r2, name, True
    inv = re.findall(r"[A-Z]{1,4}-\d{5,}", flat)
    return {"date": date, "cur": cur, "amount": float(amount), "rate": rate, "yen": yen,
            "supplier": supplier, "supplier_sure": sure, "invoices": inv}


def is_request_text(t: str) -> bool:
    return "海外送金依頼書" in t and "伝票No" in t


def parse_request(t: str, suppliers) -> dict | None:
    """楽楽精算の「海外送金依頼書」(1申請=1管理番号)。通貨・送金額・送金希望日・取引先・商品名を拾う。"""
    import difflib
    flat = re.sub(r"[ \u3000]+", " ", t)

    def field(label, pat):
        m = re.search(label + r"\s*" + pat, flat)
        return m.group(1).strip() if m else None
    cur = field("通貨種別", r"([A-Z]{3})")
    amt = field("送金額", r"(\d[\d,]*(?:\.\d+)?)")
    d = field("送金希望日", r"(\d{4}/\d{1,2}/\d{1,2})")
    if not (cur and amt and d):
        return None
    name = field("取引先名", r"(.+?)\s+品種") or ""
    product = field("品種・商品名", r"(\S+)") or ""
    memo = field("備考", r"(.+?)(?:\n|↓|$)") or ""
    supplier, sure, best = name, False, 0.0
    for sname in suppliers:
        r = difflib.SequenceMatcher(None, norm_name(name), norm_name(sname)).ratio()
        if r > best and r >= 0.78:
            best, supplier, sure = r, sname, True
    return {"date": datetime(*map(int, d.split("/"))), "cur": cur, "amount": to_num(amt),
            "supplier": supplier, "supplier_sure": sure, "product": product, "memo": memo,
            "slip": field("伝票No\\.?", r"(\d+)")}


def read_docs(pdf: Path, suppliers=(), products=None):
    """1つのPDFから、許可通知書・送金計算書・海外送金依頼書を全ページ分読む。"""
    products = products if products is not None else DEFAULT_PRODUCTS
    permits: dict[str, dict] = {}
    remits: list[dict] = []
    requests: list[dict] = []
    with pdfplumber.open(pdf) as doc:
        texts = [(p.extract_text(layout=True) or "") for p in doc.pages]
    for i, t in enumerate(texts):
        items = None
        if not t.strip():
            if i >= OCR_MAX_PAGES:
                continue
            items = ocr_items(pdf, i)
            t = "\n".join(x[2] for x in items)
        if READ_AMOUNTS and is_request_text(t):
            q = parse_request(t, suppliers)
            if q and not any(q["slip"] and q["slip"] == x["slip"] for x in requests):
                m = re.search(r"[【\[]\s*(\d{2,3}-\d{3})\s*[】\]]", pdf.name)
                q["no_hint"] = m.group(1) if m else None  # 依頼書のファイル名にある管理番号
                requests.append(q)
            continue
        if READ_AMOUNTS and is_remit_text(t, pdf.name):
            r = parse_remit(t, suppliers)
            if r and not any((r["date"], r["cur"], r["amount"]) == (x["date"], x["cur"], x["amount"]) for x in remits):
                remits.append(r)
            continue
        info = parse_items(items, suppliers) if items is not None else parse_page_text(t, suppliers)
        if info is None:
            continue
        ex = extract_extras(t, products)
        if items is not None:  # OCRでは文字の並びが崩れやすいので、位置から拾った値を優先する
            ex.update(extras_from_items(sorted(items), products, False, False))
        info.update(ex)
        if not info["shipper"] or not find_known_supplier(info["shipper"], suppliers):
            known = find_known_supplier(t, suppliers)
            if known:
                info["shipper"] = known
        permits.setdefault(info["decl"], info)
    out = []
    for info in permits.values():
        sure = True
        if info["shipper"] and suppliers:
            info["shipper"], sure = match_supplier(info["shipper"], suppliers)
        elif not info["shipper"]:
            sure = False
        info["shipper_sure"] = sure
        out.append(info)
    if not out and not remits and not requests:
        raise ValueError("許可通知書・送金計算書・海外送金依頼書が見つかりません")
    return out, remits, requests


def find_subsets(cands: list[dict], target: float, limit: int = 3) -> list[list[dict]]:
    """仕入書価格の合計が送金額に一致する組み合わせ(近い日付の許可通知書を優先)。"""
    goal = round(target * 100)
    cents = [round(c["inv_amt"] * 100) for c in cands[:20]]
    sols: list[list[dict]] = []

    def dfs(i: int, left: int, picked: list[int]):
        if len(sols) >= limit:
            return
        if left == 0 and picked:
            sols.append([cands[j] for j in picked])
            return
        for j in range(i, len(cents)):
            if cents[j] <= left:
                dfs(j + 1, left - cents[j], picked + [j])
    dfs(0, goal, [])
    return sols


def match_permits(free: list[dict], cur: str, amount: float, supplier: str, supplier_sure: bool,
                  date: datetime, window: int = 150):
    """仕入書価格の合計が amount になる許可通知書の組み合わせ。(選んだもの, 候補が複数か)"""
    cands = []
    for p in free:
        if p.get("inv_amt") is None or p.get("inv_cur") != cur:
            continue
        if p["permit"] and abs((p["permit"] - date).days) > window:
            continue
        if supplier_sure and p["shipper_sure"] and p["shipper"] and norm_name(p["shipper"]) != norm_name(supplier):
            continue
        cands.append(p)
    cands.sort(key=lambda p: abs((p["permit"] - date).days) if p["permit"] else 9999)
    sols = find_subsets(cands, amount)
    chosen = sols[0] if sols else []
    return chosen, len({tuple(sorted(id(x) for x in sol)) for sol in sols}) > 1


def build_groups(permits: list[dict], remits: list[dict], requests: list[dict] = (), window: int = 150) -> list[dict]:
    """管理番号の単位でまとめる。①海外送金依頼書(1申請=1番号) ②銀行の送金計算書 ③許可通知書だけ。
    各グループの "rem" は、台帳に書く送金情報(日付・通貨・金額・レート・円貨額)。"""
    free = list(permits)
    groups = []
    # 銀行の送金 ← 依頼書(合計が一致する組み合わせ。例: 9,800USD = 7,000 + 2,800)
    bank_of: dict[int, dict] = {}
    used_rem = set()
    reqs = [dict(q, inv_amt=q["amount"], inv_cur=q["cur"], permit=q["date"], shipper=q["supplier"],
                 shipper_sure=q["supplier_sure"], _orig=q) for q in requests]
    for rem in sorted(remits, key=lambda r: r["date"]):
        cands = [q for q in reqs if q["cur"] == rem["cur"] and id(q["_orig"]) not in bank_of
                 and abs((q["date"] - rem["date"]).days) <= 45
                 and not (rem["supplier_sure"] and q["supplier_sure"] and norm_name(q["supplier"]) != norm_name(rem["supplier"]))]
        cands.sort(key=lambda q: abs((q["date"] - rem["date"]).days))
        sols = find_subsets(cands, rem["amount"])
        if sols:
            used_rem.add(id(rem))
            for q in sols[0]:
                bank_of[id(q["_orig"])] = rem
    for q in requests:
        bank = bank_of.get(id(q))
        date = bank["date"] if bank else q["date"]
        chosen, amb = match_permits(free, q["cur"], q["amount"], q["supplier"], q["supplier_sure"], date, window)
        for pm in chosen:
            free.remove(pm)
        rate = bank["rate"] if bank else None
        yen = (round(q["amount"] * rate) if rate else None) if q["cur"] != "JPY" else int(round(q["amount"]))
        pay = {"date": date, "cur": q["cur"], "amount": q["amount"], "rate": rate, "yen": yen,
               "supplier": q["supplier"], "supplier_sure": q["supplier_sure"], "product": q["product"],
               "est": bank is None, "slip": q["slip"]}
        pay["no_hint"] = q.get("no_hint")
        groups.append({"rem": pay, "items": chosen, "ambiguous": amb, "req": q})
    # 依頼書に結びつかなかった送金計算書は、従来どおり許可通知書と直接結びつける
    for rem in sorted(remits, key=lambda r: r["date"]):
        if id(rem) in used_rem:
            continue
        chosen, amb = match_permits(free, rem["cur"], rem["amount"], rem["supplier"], rem["supplier_sure"], rem["date"], window)
        for pm in chosen:
            free.remove(pm)
        groups.append({"rem": rem, "items": chosen, "ambiguous": amb})
    groups.extend({"rem": None, "items": [pm], "ambiguous": False} for pm in free)

    def key(g):
        ds = [g["rem"]["date"]] if g["rem"] else [pm["permit"] for pm in g["items"] if pm["permit"]]
        return min(ds) if ds else datetime.max
    return sorted(groups, key=key)


def describe_groups(groups: list[dict]) -> None:
    for g in groups:
        rem, items = g["rem"], g["items"]
        if rem:
            amt = f"{rem['amount']:,.2f}" if rem["cur"] != "JPY" else f"{int(rem['amount']):,}"
            kind = ("依頼書(送金希望日・銀行の計算書なし)" if g.get("req") and rem.get("est")
                    else "依頼書+銀行の送金" if g.get("req") else "送金")
            print(f"■ {kind} {rem['date']:%Y/%m/%d}  {rem['cur']} {amt}  {rem['supplier'] or '(受取人不明)'}"
                  + (f"  → 許可通知書 {len(items)} 件" if items else "  → 金額の合う許可通知書が見つかりません(要確認)")
                  + ("  ※合う組み合わせが複数あります(要確認)" if g["ambiguous"] else "")
                  + (f"  管理番号(依頼書のファイル名) {rem['no_hint']}" if rem.get("no_hint") else ""))
        elif READ_AMOUNTS:
            print("■ 送金計算書が見つかりません(送金の欄は空欄・黄色になります)")
        for p in sorted(items, key=lambda x: x['permit'] or datetime.max):
            d = f"{p['permit']:%Y/%m/%d}" if p["permit"] else "(許可日なし)"
            flag = "" if p["shipper_sure"] else "  ※仕出人要確認"
            print(f"    申告番号 {p['decl']}  許可日 {d}  数量 {p['qty'] or '不明'}  {p['product'] or ''}  {p['shipper']}{flag}")
    print()


DEFAULT_SUPPLIERS = ["TAEKWANG INDUSTRIAL CO.,LTD.", "CHUEN HUAH CHEMICAL CO.,LTD."]


def load_suppliers(root: Path) -> list[str]:
    """仕入先名の一覧。正式名(仕入先名.txt)を先に、台帳に出てきた表記ゆれを後ろに並べる。
    同じ綴りなら先頭(正式名)が優先されるので、OCRの崩れた表記が台帳に残らない。"""
    f = root / "仕入先名.txt"
    if not f.exists():  # 初回は見本を作る。社名を足したいときは1行1社で追記する
        root.mkdir(parents=True, exist_ok=True)
        f.write_text("\n".join(DEFAULT_SUPPLIERS) + "\n", encoding="utf8")
    official = [l.strip() for l in f.read_text(encoding="utf8").splitlines() if l.strip()]
    seen = {norm_name(n) for n in official}
    extra = []
    for lg in sorted(root.glob("第*期/輸入明細一覧_*.xlsx")):
        wb = openpyxl.load_workbook(lg, read_only=True)
        for ws, col in ((wb[SHEET_SEND], 4), (wb[SHEET_NOSEND], 3)):
            for r in ws.iter_rows(min_row=4, min_col=col, max_col=col, values_only=True):
                n = str(r[0]).strip() if r[0] else ""
                if n and norm_name(n) not in seen:
                    seen.add(norm_name(n))
                    extra.append(n)
    return official + extra


# ---------- 台帳 ----------
def new_ledger(path: Path, period: int):
    wb = openpyxl.load_workbook(open_template(TPL_LEDGER, _TPL_LEDGER_B64))
    for ws in wb:
        # 見本データを消し、見出しだけ残す(「例」行は送金ありのみ残す)
        first = 5 if ws.title == SHEET_SEND else 4
        for row in ws.iter_rows(min_row=first, max_row=ws.max_row):
            for c in row:
                c.value = None
    wb[SHEET_SEND]["B1"] = period_label(period)
    wb[SHEET_NOSEND]["A1"] = period_label(period)
    wb.save(path)


def next_no(ws, col: int, prefix: str, period: int, reserved=()) -> str:
    """続きの管理番号。依頼書のファイル名で予約された番号(reserved)も飛ばす。"""
    pat = re.compile(rf"^{re.escape(prefix)}{period}-(\d+)$")
    nums = [int(m.group(1)) for r in range(5 if ws.title == SHEET_SEND else 4, ws.max_row + 1)
            if (v := ws.cell(r, col).value) and (m := pat.match(str(v)))]
    nums += [int(m.group(1)) for v in reserved if (m := pat.match(str(v)))]
    return f"{prefix}{period}-{max(nums, default=0) + 1:03d}"


def next_row(ws, key_col: int) -> int:
    r = ws.max_row
    while r >= 4 and ws.cell(r, key_col).value is None and ws.cell(r, key_col + 1).value is None:
        r -= 1
    return max(r + 1, 5 if ws.title == SHEET_SEND else 4)


def find_shared_no(ws, nosend: bool, info: dict) -> str | None:
    """同じ仕出人・同じ許可月の行が台帳にあれば、その管理番号を返す(まとめ対象の仕出人のみ)。"""
    key = norm_name(info["shipper"])
    if not info["permit"] or key not in {norm_name(n) for n in MERGE_MONTHLY}:
        return None
    no_col, ship_col, date_col = (2, 10, 11) if nosend else (3, 14, 15)
    for r in range(5 if not nosend else 4, ws.max_row + 1):
        d, no, ship = ws.cell(r, date_col).value, ws.cell(r, no_col).value, ws.cell(r, ship_col).value
        if (no and ship and isinstance(d, datetime) and norm_name(str(ship)) == key
                and (d.year, d.month) == (info["permit"].year, info["permit"].month)):
            return str(no)
    return None


def add_send(ws, no, office, info):
    r = next_row(ws, 3)
    vals = {2: office, 3: no, 13: info["decl"], 14: info["shipper"], 4: info["shipper"],
            15: info["permit"]}
    for c, v in vals.items():
        ws.cell(r, c).value = v
    return r


def add_nosend(ws, no, office, info):
    r = next_row(ws, 2)
    vals = {1: office, 2: no, 3: info["shipper"], 9: info["decl"], 10: info["shipper"],
            11: info["permit"]}
    for c, v in vals.items():
        ws.cell(r, c).value = v
    return r


# ---------- 表紙 ----------
def cover_data(ws, no: str, nosend: bool) -> dict:
    """台帳から、同じ管理番号の申告番号(全件)・許可日(最も早い日)・仕出人を集める。"""
    c_no, c_decl, c_date, c_ship = (2, 9, 11, 10) if nosend else (3, 13, 15, 14)
    decls, dates, shipper = [], [], ""
    for r in range(4 if nosend else 5, ws.max_row + 1):  # 送金ありシートの4行目は「例」
        if str(ws.cell(r, c_no).value or "") != no:
            continue
        d = str(ws.cell(r, c_decl).value or "").replace(" ", "")
        if d and d not in [x[1] for x in decls]:
            dt = ws.cell(r, c_date).value
            decls.append((dt if isinstance(dt, datetime) else datetime.max, d))
        if isinstance(ws.cell(r, c_date).value, datetime):
            dates.append(ws.cell(r, c_date).value)
        shipper = shipper or str(ws.cell(r, c_ship).value or "")
    decls.sort()
    return {"decls": [d for _, d in decls], "permit": min(dates) if dates else None, "shipper": shipper}


def load_cover_sheet(nosend: bool):
    """表紙のひな形から、海外送金あり/なしのシートを探して返す(他のシートは消す)。
    templates のひな形に目的のシートが無ければ、埋め込みのひな形を使う。"""
    want = "海外送金なし" if nosend else "海外送金あり"
    for src in (TPL_COVER, None):
        try:
            wb = openpyxl.load_workbook(open_template(TPL_COVER, _TPL_COVER_B64) if src else
                                        __import__("io").BytesIO(__import__("base64").b64decode("".join(_TPL_COVER_B64.split()))))
        except Exception:
            continue
        for sh in wb.worksheets:
            if str(sh["A1"].value or "").strip().startswith(want):
                for other in list(wb.worksheets):
                    if other is not sh:
                        wb.remove(other)
                return wb, sh
    raise ValueError("表紙のひな形に「" + want + "」のシートが見つかりません")


def make_cover(out: Path, no: str, info: dict, nosend: str | None, remit_yen=None, many=False):
    """表紙を作る。info["decls"] があれば、申告番号を1件ずつ別の行に並べる。"""
    wb, ws = load_cover_sheet(nosend is not None)
    if nosend is None:
        ws["E16"] = None  # 見本の「乙仲 No.」を消す
        ymd = ("E10", "I10", "K10")
        if remit_yen:
            ws["J48"] = int(remit_yen)
            ws["J48"].number_format = "#,##0"
    else:
        ymd = ("E10", "H10", "J10")
        if nosend in ("着払", "無償"):
            ws["H1" if nosend == "着払" else "J1"].fill = HILITE
    decls = info.get("decls") or [info["decl"].replace(" ", "")]
    ws["E3"] = no
    if "E5:Q5" not in {str(m) for m in ws.merged_cells.ranges}:
        ws.merge_cells("E5:Q5")
    from copy import copy
    from openpyxl.styles import Alignment
    ws["E5"] = "\n".join(decls)
    al = copy(ws["E5"].alignment)
    ws["E5"].alignment = Alignment(horizontal=al.horizontal or "left", vertical="center", wrap_text=True)
    size = ws["E5"].font.sz or 11
    ws.row_dimensions[5].height = max(30, len(decls) * (size * 1.5 + 2))
    ws["E7"] = info["shipper"]
    if info["permit"]:
        for cell, v in zip(ymd, (info["permit"].year, info["permit"].month, info["permit"].day)):
            ws[cell] = v
    wb.save(out)


# ---------- main ----------
def collect(paths):
    for p in map(Path, paths):
        if p.is_dir():
            yield from sorted(p.rglob("*.pdf"))
        else:
            yield p


SETTINGS = Path(__file__).resolve().parent / "settings.json"


def log_error(root: Path, title: str) -> None:
    import traceback
    try:
        root.mkdir(parents=True, exist_ok=True)
        with open(root / "実行ログ.txt", "a", encoding="utf8") as f:
            f.write(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] {title}\n{traceback.format_exc()}\n")
    except OSError:
        pass


def acquire_lock(root: Path):
    """同時に2つ動かして台帳が壊れるのを防ぐ。"""
    import os
    import time
    root.mkdir(parents=True, exist_ok=True)
    path = root / ".実行中"
    if path.exists() and time.time() - path.stat().st_mtime > 1800:  # 30分以上前の残りは無効
        path.unlink()
    try:
        os.close(os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
    except FileExistsError:
        raise SystemExit("別の実行がまだ動いています。黒い画面が残っていないか確認して、終わってからやり直してください。"
                         f"\n(動いていないのに出る場合は、{path} を削除してください)")
    return path


def release_lock(path: Path) -> None:
    try:
        path.unlink()
    except OSError:
        pass


def run_all(a, root: Path, paths, confirm=None) -> int:
    """PDFを読み、送金ごとにまとめて表示し、(確認のうえ)台帳と表紙に保存する。保存した件数を返す。"""
    suppliers, products = load_suppliers(root), load_products(root)
    permits, remits, requests, skipped, seen, errors = [], [], [], 0, set(), []
    for pdf in collect(paths):
        try:
            ps, rs, qs = read_docs(pdf, suppliers, products)
        except ValueError:
            skipped += 1  # 請求書・到着案内など、許可通知書・送金計算書ではないPDF
            continue
        except Exception as e:  # 想定外の失敗は隠さず、ログに残す
            errors.append(f"{pdf.name}: {type(e).__name__}: {e}")
            log_error(root, f"読み取りエラー {pdf}")
            continue
        for p in ps:
            if p["decl"] not in seen:
                seen.add(p["decl"])
                p["pdf"] = pdf
                permits.append(p)
        for r in rs:
            if not any((r["date"], r["cur"], r["amount"]) == (x["date"], x["cur"], x["amount"]) for x in remits):
                r["pdf"] = pdf
                remits.append(r)
        for q in qs:
            if not any(q["slip"] and q["slip"] == x["slip"] for x in requests):
                q["pdf"] = pdf
                requests.append(q)
    print(f"許可通知書 {len(permits)} 件、送金計算書 {len(remits)} 件、海外送金依頼書 {len(requests)} 件を読み取りました"
          f"(それ以外のPDF {skipped} 件はスキップ)\n")
    if errors:
        print(f"※読み取りに失敗したPDFが {len(errors)} 件あります(詳細は {root / '実行ログ.txt'}):")
        for e in errors[:10]:
            print("   ", e)
        print()
    if not permits and not remits and not requests:
        return 0
    if a.nosend:
        groups = [{"rem": None, "items": [p], "ambiguous": False}
                  for p in sorted(permits, key=lambda x: x["permit"] or datetime.max)]
    else:
        groups = build_groups(permits, remits, requests)
    describe_groups(groups)
    a.reserved = {g["rem"]["no_hint"] for g in groups if g["rem"] and g["rem"].get("no_hint")}
    if a.dry_run:
        return len(groups)
    if confirm and not confirm(len(groups)):
        return 0
    lock = acquire_lock(root)
    try:
        n = sum(save_group(a, root, g) for g in groups)
    except PermissionError as e:
        raise SystemExit(f"\n保存できません: {e.filename}\n台帳や表紙を Excel で開いている場合は、閉じてからもう一度実行してください。")
    finally:
        release_lock(lock)
    return n


def gui() -> int:
    """引数なしで起動したとき(ダブルクリック用)。画面でフォルダ・事業所名・種別を選ぶ。"""
    print(f"ツールの版: {VERSION}\n")
    import json
    import os
    import tkinter as tk
    from argparse import Namespace
    from tkinter import filedialog, messagebox, simpledialog

    root_win = tk.Tk()
    root_win.withdraw()
    root_win.attributes("-topmost", True)
    conf = {}
    if SETTINGS.exists():
        conf = json.loads(SETTINGS.read_text(encoding="utf8"))

    # まず場所の貼り付け欄を出す(ドライブの隠しフォルダは選択画面から選べないことがあるため)
    src = simpledialog.askstring("フォルダの場所",
                                 "許可通知書のフォルダの場所を貼り付けてください。\n"
                                 "(エクスプローラーのアドレス欄をコピー → ここで Ctrl+V)\n\n"
                                 "空欄のまま OK を押すと、フォルダ選択画面が開きます。",
                                 initialvalue=conf.get("last_dir", ""), parent=root_win)
    if src is None:
        return 1
    src = src.strip().strip('"')
    if not src:
        src = filedialog.askdirectory(title="許可通知書の入っているフォルダを選んでください",
                                      initialdir=str(Path.home()))
    if not src or not Path(src).is_dir():
        messagebox.showerror("フォルダが見つかりません", f"次の場所が見つかりません:\n{src}")
        return 1
    office = simpledialog.askstring("事業所名", "台帳の「事業所名」に入れる名前(例: 林六／東京)",
                                    initialvalue=conf.get("office", ""), parent=root_win)
    if office is None:
        return 1
    send = messagebox.askyesnocancel("種別", "海外送金ありの分ですか?\n\n"
                                     "はい … 海外送金あり\nいいえ … 海外送金なし(着払・無償・乙仲)")
    if send is None:
        return 1
    nosend = None
    if not send:
        nosend = simpledialog.askstring("海外送金なしの種別", "着払 / 無償 / 乙仲 のどれかを入力", initialvalue="着払",
                                        parent=root_win)
        if nosend not in ("着払", "無償", "乙仲"):
            messagebox.showerror("入力エラー", "「着払」「無償」「乙仲」のどれかを入力してください")
            return 1
    SETTINGS.write_text(json.dumps({"office": office, "last_dir": src}, ensure_ascii=False), encoding="utf8")

    out = desktop() / "輸入事後調査"
    a = Namespace(office=office, nosend=nosend, period=None, dry_run=False, out=out)
    print(f"読み取り中: {src}\n(スキャンPDFが多いと時間がかかります)\n")

    def confirm(n):
        return messagebox.askyesno("保存の確認", f"{n} 件の送金・許可通知書を、台帳と表紙に保存しますか?\n\n"
                                   "黒い画面の内容を確認してください。\n黄色になる欄は、保存後に台帳で確認してください。")
    try:
        n = run_all(a, out, [src], confirm)
    except SystemExit as e:
        print(e)
        messagebox.showerror("保存できませんでした", str(e))
        return 1
    except Exception:
        import traceback
        print(traceback.format_exc())
        log_error(out, "実行中のエラー")
        messagebox.showerror("エラー", f"途中で失敗しました。\n黒い画面の内容と、{out}\\実行ログ.txt を確認してください。")
        return 1
    if not n:
        messagebox.showinfo("結果", "保存したものはありません。\n黒い画面の内容(件数・エラー)を確認してください。")
        return 1
    period = fiscal_period(datetime.now())
    ledger = out / f"第{period}期" / f"輸入明細一覧_{period}期.xlsx"
    messagebox.showinfo("完了", f"{n} 件を保存しました。\n\n{ledger}\n\nExcel(台帳)を開きます。")
    for target in (ledger, out):
        try:
            os.startfile(target)
        except Exception:
            pass
    return 0


def main():
    if len(sys.argv) == 1:
        return gui()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf", nargs="+", help="輸入許可通知書の PDF またはそのフォルダ")
    ap.add_argument("--office", default="", help="事業所名/担当(例: 林六／東京)")
    ap.add_argument("--nosend", choices=["着払", "無償", "乙仲"], help="海外送金なしの場合の種別")
    ap.add_argument("--period", type=int, help="期(省略時は今日の日付から判定)")
    ap.add_argument("--dry-run", action="store_true", help="読み取り結果を表示するだけで保存しない(検証用)")
    ap.add_argument("--out", type=Path, help="出力先(省略時はデスクトップ/輸入事後調査)")
    a = ap.parse_args()

    root = a.out or desktop() / "輸入事後調査"
    ok = run_all(a, root, a.pdf)
    print(f"{ok} 件を確認しました(保存なし)" if a.dry_run else f"{ok} 件を {root} に保存しました")
    return 0 if ok else 1


def process(a, root: Path, pdf: Path, info: dict) -> int:
    """許可通知書1件分を台帳に追記し、表紙を作る。成功なら 1。"""
    if a.dry_run:
        d = f"{info['permit']:%Y/%m/%d}" if info["permit"] else "(許可日なし)"
        flag = "" if info["shipper_sure"] else "  ※仕出人要確認"
        print(f"[確認] {pdf.name}: 申告番号 {info['decl']}  許可日 {d}  仕出人 {info['shipper']}{flag}")
        return 1
    period = a.period or fiscal_period(datetime.now())  # 期は処理日基準(許可日が前期のこともある)
    folder = root / f"第{period}期"
    (folder / "表紙").mkdir(parents=True, exist_ok=True)
    ledger = folder / f"輸入明細一覧_{period}期.xlsx"
    if not ledger.exists():
        new_ledger(ledger, period)
    wb = openpyxl.load_workbook(ledger)
    for sh in wb.worksheets:
        sh.sheet_state = "visible"
    ws = wb[SHEET_NOSEND] if a.nosend else wb[SHEET_SEND]
    col = 9 if a.nosend else 13
    want = info["decl"].replace(" ", "")
    dup = [r for r in range(4, ws.max_row + 1) if str(ws.cell(r, col).value or "").replace(" ", "") == want]
    if dup:  # 同じ申告番号の二重登録を防ぐ
        print(f"[SKIP] {pdf.name}: 申告番号 {info['decl']} は台帳の {dup[0]} 行目に登録済み", file=sys.stderr)
        return 0
    # 管理番号をまとめる: ①同じフォルダの許可通知書 ②同月の同じ仕出人(MERGE_MONTHLY)
    fmap_path = folder / "フォルダ別管理番号.json"
    fmap = json.loads(fmap_path.read_text(encoding="utf8")) if fmap_path.exists() else {}
    fkey = f"{'なし' if a.nosend else 'あり'}|{pdf.parent}"
    used = {str(ws.cell(r, 2 if a.nosend else 3).value) for r in range(4, ws.max_row + 1)}
    shared = fmap.get(fkey) if fmap.get(fkey) in used else None
    by_folder = bool(shared)
    shared = shared or find_shared_no(ws, bool(a.nosend), info)
    if a.nosend:
        no = shared or next_no(ws, 2, "経", period)
        row = add_nosend(ws, no, a.office, info)
    else:
        no = shared or next_no(ws, 3, "", period)
        row = add_send(ws, no, a.office, info)
    if not info["shipper_sure"]:
        for c in ((row, 3), (row, 10)) if a.nosend else ((row, 4), (row, 14)):
            ws.cell(*c).fill = HILITE
        print(f"      仕出人『{info['shipper']}』は照合できませんでした(黄色の欄を要確認)")
    wb.save(ledger)
    fmap[fkey] = no
    fmap_path.write_text(json.dumps(fmap, ensure_ascii=False, indent=1), encoding="utf8")
    cd = cover_data(ws, no, bool(a.nosend))  # 台帳の同じ管理番号の申告番号を、全件ならべる
    make_cover(folder / "表紙" / f"表紙_{no}.xlsx", no, cd, a.nosend)
    d = f"{info['permit']:%Y/%m/%d}" if info["permit"] else "(許可日は手入力)"
    note = ("  ※同じフォルダのため管理番号をまとめました" if by_folder
            else "  ※同月の同じ仕出人のため管理番号をまとめました") if shared else ""
    print(f"[OK] {pdf.name} -> {no}  申告番号 {info['decl']}  許可日 {d}{note}")
    return 1


# ---------- 追加: 送金グループの保存 ----------
NOFILL = PatternFill(fill_type=None)


def clear_fill(ws, row, c1=3, c2=15):
    for c in range(c1, c2 + 1):
        ws.cell(row, c).fill = NOFILL


def save_group(a, root: Path, g: dict) -> int:
    """送金1件(または許可通知書1件)を台帳に書き、表紙を作る。"""
    items, rem = sorted(g["items"], key=lambda x: x["permit"] or datetime.max), g["rem"]
    if a.nosend:  # 海外送金なしは、送金計算書がないので従来どおり1件ずつ
        return sum(process(a, root, p["pdf"], p) for p in items)
    period = a.period or fiscal_period(datetime.now())
    folder = root / f"第{period}期"
    (folder / "表紙").mkdir(parents=True, exist_ok=True)
    ledger = folder / f"輸入明細一覧_{period}期.xlsx"
    if not ledger.exists():
        new_ledger(ledger, period)
    wb = openpyxl.load_workbook(ledger)
    for sh in wb.worksheets:  # 非表示になっているシートも表示する
        sh.sheet_state = "visible"
    wb.active = 0
    ws = wb[SHEET_SEND]

    # 前回、許可通知書が結びつかず「送金だけ」で作った行は、今回結びついたら削除する(二重にならないように)
    if rem and items:
        for r in range(ws.max_row, 4, -1):
            if (not ws.cell(r, 13).value and ws.cell(r, 6).value == rem["date"]
                    and ws.cell(r, 7).value == ("円" if rem["cur"] == "JPY" else rem["cur"])):
                ws.delete_rows(r)

    # 既に台帳にある行(申告番号が同じ)は、その行を更新する
    rows = {}
    for r in range(5, ws.max_row + 1):
        v = str(ws.cell(r, 13).value or "").replace(" ", "")
        if v:
            rows[v] = r
    exist = [rows[p["decl"].replace(" ", "")] for p in items if p["decl"].replace(" ", "") in rows]
    nos = []
    for r in exist:
        n = ws.cell(r, 3).value
        if n and n not in nos:
            nos.append(str(n))
    ambiguous = g["ambiguous"]
    # 送金計算書が見つからない許可通知書だけは、従来のまとめ方(同じフォルダ・同月のTAEKWANG)を使う
    fallback = not rem and bool(items) and not nos
    fmap_path = folder / "フォルダ別管理番号.json"
    fmap = json.loads(fmap_path.read_text(encoding="utf8")) if fmap_path.exists() else {}
    fkey = f"あり|{items[0]['pdf'].parent}" if items else ""
    shared = None
    if nos:
        no = nos[0]
    else:
        if fallback:
            used = {str(ws.cell(r, 3).value) for r in range(5, ws.max_row + 1)}
            shared = fmap.get(fkey) if fmap.get(fkey) in used else find_shared_no(ws, False, items[0])
        hint = (rem or {}).get("no_hint")
        taken = {str(ws.cell(r, 3).value) for r in range(5, ws.max_row + 1)}
        if hint and hint.startswith(f"{period}-") and hint not in taken:
            no = hint  # 依頼書のファイル名にある管理番号(まだ台帳で使われていない場合)
        else:
            no = shared or next_no(ws, 3, "", period, getattr(a, "reserved", ()))
    if len(nos) > 1:
        print(f"      管理番号 {', '.join(nos)} を {no} にまとめました")

    # 行の中身
    cur = rem["cur"] if rem else None
    n_items = len(items)
    yens = []
    for p in items:
        if rem and cur != "JPY" and rem["rate"]:
            yens.append(round(p["inv_amt"] * rem["rate"]))
        elif rem and cur == "JPY":
            yens.append(int(round(p["inv_amt"])))
        else:
            yens.append(None)
    if rem and cur != "JPY" and rem["rate"] and rem["yen"] and yens:
        yens[-1] += rem["yen"] - sum(yens)  # 端数は最後の行で合わせる

    lines = []
    for k, p in enumerate(items):
        lines.append((p, yens[k] if k < len(yens) else None))
    if rem and not items:
        lines.append((None, rem["yen"]))

    def put(row, col, val):  # 空欄だけ埋める(手で直した内容は上書きしない)
        if val is not None and val != "" and ws.cell(row, col).value in (None, ""):
            ws.cell(row, col).value = val

    for p, yen in lines:
        row = rows.get(p["decl"].replace(" ", "")) if p else None
        if row is None:
            row = next_row(ws, 3)
        clear_fill(ws, row)
        put(row, 2, a.office)
        ws.cell(row, 3).value = no  # 管理番号は常に最新(同じ送金は同じ番号)
        put(row, 4, (p["shipper"] if p else rem["supplier"]) or None)
        if p:
            put(row, 14, p["shipper"])
            put(row, 5, p["product"] or (rem or {}).get("product"))
            put(row, 9, p["qty"])
            put(row, 10, p["unit"])
            put(row, 13, p["decl"])
            put(row, 15, p["permit"])
        if rem:
            put(row, 6, rem["date"])
            put(row, 7, "円" if cur == "JPY" else cur)
            if cur != "JPY":
                put(row, 8, p["inv_amt"] if p else rem["amount"])
            put(row, 12, yen)
            if p and cur == "JPY" and p["qty"] and yen:
                put(row, 11, round(yen / p["qty"], 3))
        # 黄色: 自動で埋まらなかった欄、確認が必要な欄
        check = (4, 5, 6, 7, 9, 10, 12, 13, 14, 15) if READ_AMOUNTS else (4, 5, 9, 10, 13, 14, 15)
        warn = [c for c in check if ws.cell(row, c).value in (None, "")]
        if p and not p["shipper_sure"]:
            warn += [4, 14]
        if rem and not p and not rem["supplier_sure"]:
            warn += [4]
        if ambiguous:
            warn += [3]
        if rem and rem.get("est"):
            warn += [6]  # 送金日が依頼書の「送金希望日」(銀行の計算書で確認できていない)
        for c in warn:
            ws.cell(row, c).fill = HILITE
    wb.save(ledger)

    if fallback:
        fmap[fkey] = no
        fmap_path.write_text(json.dumps(fmap, ensure_ascii=False, indent=1), encoding="utf8")

    # 表紙(管理番号ごとに1枚。同じ管理番号の申告番号は、台帳から全件を1件ずつ並べる)
    cd = cover_data(ws, no, False)
    if not cd["decls"] and rem:
        cd = {"decls": [""], "permit": rem["date"], "shipper": rem["supplier"]}
    make_cover(folder / "表紙" / f"表紙_{no}.xlsx", no, cd, None, rem["yen"] if (rem and READ_AMOUNTS) else None)
    d = rem["date"].strftime("%Y/%m/%d") if rem else "送金なし"
    if READ_AMOUNTS:
        print(f"[OK] {no}  送金 {d}  許可通知書 {len(items)} 件" + ("  ※黄色の欄を確認" if (not rem or ambiguous) else ""))
    else:
        total = sum(1 for r in range(5, ws.max_row + 1) if str(ws.cell(r, 3).value or "") == no)
        print(f"[OK] {no}  {', '.join(p['decl'] for p in items)}  (この管理番号の申告番号: {total} 件)")
    return 1



# ---------- ひな形(台帳・表紙)を埋め込み。templates フォルダが無くても動く ----------
_TPL_LEDGER_B64 = """
UEsDBBQABgAIAAAAIQAhjEY6cwEAAIwFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADEVMluwjAQvVfqP0S+VomBQ1VVBA5dji0S9ANMPCEW
iW15Bgp/34lZVFUsQiD1kiix5232TH+4aupkCQGNs7noZh2RgC2cNnaWi6/Je/okEiRltaqdhVysAcVwcH/Xn6w9YMLVFnNREfln
KbGooFGYOQ+WV0oXGkX8GWbSq2KuZiB7nc6jLJwlsJRSiyEG/Vco1aKm5G3FvzdKpsaK5GWzr6XKhfK+NoUiFiqXVv8hSV1ZmgK0
KxYNQ2foAyiNFQA1deaDYcYwBiI2hkIe5AxQ42WkW1cZV0ZhWBmPD2z9CEO7ctzVtu6TjyMYDclIBfpQDXuXq1p+uzCfOjfPToNc
Gk2MKGuUsTvdJ/jjZpTx1b2xkNZfBL5QR++fdBDfdZDxeX0UEeaMcaR1DXjr44+g55grFUCPibtodnMBv7FP6eDWHgXnkadHgMtT
2LVqW516BoJABvbNeujS7xl59FwdO7SzTYM+wC3jLB38AAAA//8DAFBLAwQUAAYACAAAACEAtVUwI/QAAABMAgAACwAIAl9yZWxz
Ly5yZWxzIKIEAiigAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AKySTU/DMAyG70j8h8j31d2QEEJLd0FIuyFUfoBJ3A+1jaMkG92/JxwQVBqDA0d/vX78ytvdPI3qyCH24jSsixIUOyO2d62Gl/px
dQcqJnKWRnGs4cQRdtX11faZR0p5KHa9jyqruKihS8nfI0bT8USxEM8uVxoJE6UchhY9mYFaxk1Z3mL4rgHVQlPtrYawtzeg6pPP
m3/XlqbpDT+IOUzs0pkVyHNiZ9mufMhsIfX5GlVTaDlpsGKecjoieV9kbMDzRJu/E/18LU6cyFIiNBL4Ms9HxyWg9X9atDTxy515
xDcJw6vI8MmCix+o3gEAAP//AwBQSwMEFAAGAAgAAAAhACwcr3ySBAAAcAsAAA8AAAB4bC93b3JrYm9vay54bWysVt1r21YUfx/s
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
bfOze6sV/7dwb7/ePIRf62v+42a5nS2CtljP1vPdbt7MrT4dLz8cw6/0zZ7q/zBf7ZpfZ/h4cxPu2v+Pzd2EVf/j+PDt9uT28vZ4
94/r/zyGPyoRtMfV5efwb83Pr3+8fvjnjfxQYDhc3dxdh98HvHy4vvl6fvr58uuHgL09np78cbwLtFx+vri9DpcLd/qsIfbu7Ye2
VOMi37SrefH85sMH/Ov/ufxy+3//tf3fvzw/6+3Pz9jj7M+bu9/bjL74/wAAAP//AwBQSwMEFAAGAAgAAAAhAL6YT3RmHwAAscUA
ABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0Mi54bWy0XV1zG8txfU9V/gMLD3mTiMUXQYWUSyS4krCV1C3f6+QZIkEJdUmCAaGPa5f/
u2d2Tu/2mW7QLEeTciz56Gzv7Nme2dOzAPrsTz/u746+rXdPm+3D+aB6PRwcrR+utzebh8/ng7/8Vr+aD46e9quHm9Xd9mF9Pvhj
/TT409t//7ez79vd709f1uv9UYjw8HQ++LLfP745Pn66/rK+Xz293j6uH8K/3G5396t9+J+7z8dPj7v16qY96P7ueDQczo7vV5uH
QYrwZveSGNvb2831erG9/nq/ftinILv13Wofxv/0ZfP4JNHur18S7n61+/3r46vr7f1jCPFpc7fZ/9EGHRzdX7/5+Plhu1t9ugvX
/aOarK6PfuzCf0bh/8dymhY3Z7rfXO+2T9vb/esQ+TiN2V7+6fHp8eq6i2Sv/0Vhqsnxbv1tE29gH2r0rw2pmnaxRn2w8b8YbNYF
i3Lt3nzd3JwP/lZdXYym9Wj0ajx+V7+ajGanry4mi/Gri4vZ5GI8HI/n05O/D96etXnyy+7t2ePq8/rX9f4vj7/sjm43+9+2vwQg
5Org+O3Zcce62YSEiCIc7da354N31Zummo4jp6X8z2b9/Un9/Wi/+vTr+m59vV+HQVWDo79ut/e/Xq/ivZ6HSdD9z/+OCXyXwJjz
n7bb32Owj+GwYRxmGySed3W933xbX67vAvuiOgnz5v/aocS/d0ONh8qw9Zjqdp6EK7xZ366+3u0vt3f/u7nZfzkfnA4E+/P2+4f1
5vOXfRjw/PXJNIgas/LNzR+L9dN1mA5hRK8n8VzX27twseG/j+43cVqHbF79aP/8noJW49fjGOBp/0e84kD4tH7a15sYenB0/fVp
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
"""
_TPL_COVER_B64 = """
UEsDBBQABgAIAAAAIQAhjEY6cwEAAIwFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADEVMluwjAQvVfqP0S+VomBQ1VVBA5dji0S9ANMPCEW
iW15Bgp/34lZVFUsQiD1kiix5232TH+4aupkCQGNs7noZh2RgC2cNnaWi6/Je/okEiRltaqdhVysAcVwcH/Xn6w9YMLVFnNREfln
KbGooFGYOQ+WV0oXGkX8GWbSq2KuZiB7nc6jLJwlsJRSiyEG/Vco1aKm5G3FvzdKpsaK5GWzr6XKhfK+NoUiFiqXVv8hSV1ZmgK0
KxYNQ2foAyiNFQA1deaDYcYwBiI2hkIe5AxQ42WkW1cZV0ZhWBmPD2z9CEO7ctzVtu6TjyMYDclIBfpQDXuXq1p+uzCfOjfPToNc
Gk2MKGuUsTvdJ/jjZpTx1b2xkNZfBL5QR++fdBDfdZDxeX0UEeaMcaR1DXjr44+g55grFUCPibtodnMBv7FP6eDWHgXnkadHgMtT
2LVqW516BoJABvbNeujS7xl59FwdO7SzTYM+wC3jLB38AAAA//8DAFBLAwQUAAYACAAAACEAtVUwI/QAAABMAgAACwAIAl9yZWxz
Ly5yZWxzIKIEAiigAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AKySTU/DMAyG70j8h8j31d2QEEJLd0FIuyFUfoBJ3A+1jaMkG92/JxwQVBqDA0d/vX78ytvdPI3qyCH24jSsixIUOyO2d62Gl/px
dQcqJnKWRnGs4cQRdtX11faZR0p5KHa9jyqruKihS8nfI0bT8USxEM8uVxoJE6UchhY9mYFaxk1Z3mL4rgHVQlPtrYawtzeg6pPP
m3/XlqbpDT+IOUzs0pkVyHNiZ9mufMhsIfX5GlVTaDlpsGKecjoieV9kbMDzRJu/E/18LU6cyFIiNBL4Ms9HxyWg9X9atDTxy515
xDcJw6vI8MmCix+o3gEAAP//AwBQSwMEFAAGAAgAAAAhAHVd10+4AwAABgkAAA8AAAB4bC93b3JrYm9vay54bWysVd1u2zYUvh+w
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
ABInp5EiCAAAZiQAABgAAAAAAAAAAAAAAAAA8AwAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbFBLAQItABQABgAIAAAAIQAp7spi
cQcAACkgAAAYAAAAAAAAAAAAAAAAAEgVAAB4bC93b3Jrc2hlZXRzL3NoZWV0Mi54bWxQSwECLQAUAAYACAAAACEA7VXUyMUHAAAV
IgAAEwAAAAAAAAAAAAAAAADvHAAAeGwvdGhlbWUvdGhlbWUxLnhtbFBLAQItABQABgAIAAAAIQCRID7C0wQAACIdAAANAAAAAAAA
AAAAAAAAAOUkAAB4bC9zdHlsZXMueG1sUEsBAi0AFAAGAAgAAAAhAC6G7ubYBQAAVxoAABQAAAAAAAAAAAAAAAAA4ykAAHhsL3No
YXJlZFN0cmluZ3MueG1sUEsBAi0AFAAGAAgAAAAhADttMkvBAAAAQgEAACMAAAAAAAAAAAAAAAAA7S8AAHhsL3dvcmtzaGVldHMv
X3JlbHMvc2hlZXQxLnhtbC5yZWxzUEsBAi0AFAAGAAgAAAAhABPELBPCAAAAQgEAACMAAAAAAAAAAAAAAAAA7zAAAHhsL3dvcmtz
aGVldHMvX3JlbHMvc2hlZXQyLnhtbC5yZWxzUEsBAi0AFAAGAAgAAAAhAEf9F0kMAwAAzAwAACcAAAAAAAAAAAAAAAAA8jEAAHhs
L3ByaW50ZXJTZXR0aW5ncy9wcmludGVyU2V0dGluZ3MxLmJpblBLAQItABQABgAIAAAAIQBH/RdJDAMAAMwMAAAnAAAAAAAAAAAA
AAAAAEM1AAB4bC9wcmludGVyU2V0dGluZ3MvcHJpbnRlclNldHRpbmdzMi5iaW5QSwECLQAUAAYACAAAACEAfxbCvYYBAACeAgAA
EQAAAAAAAAAAAAAAAACUOAAAZG9jUHJvcHMvY29yZS54bWxQSwECLQAUAAYACAAAACEAUI1qGrMBAAA6AwAAEAAAAAAAAAAAAAAA
AABROwAAZG9jUHJvcHMvYXBwLnhtbFBLBQYAAAAADwAPABIEAAA6PgAAAAA=
"""


if __name__ == "__main__":
    sys.exit(main())
