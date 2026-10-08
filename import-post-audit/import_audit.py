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
VERSION = "2026-10-08 f (送金計算書の読み取り・送金ごとの管理番号)"
SHEET_SEND = "海外送金あり"
SHEET_NOSEND = "海外送金なし（乙仲・無償・着払）"
HILITE = PatternFill("solid", fgColor="FFFF00")

# 同じ月(輸入許可日の年月)に複数あっても、管理番号を1つにまとめる仕出人
MERGE_MONTHLY = ["TAEKWANG INDUSTRIAL CO.,LTD."]


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


def extract_extras(text: str, products: dict) -> dict:
    """許可通知書から仕入書価格(通貨・金額)・数量・商品名を拾う。"""
    flat = re.sub(r"[ 　]+", " ", text)
    out = {"inv_cur": None, "inv_amt": None, "qty": None, "unit": "", "product": ""}
    m = re.search(r"(?:CIF|FOB|CFR|C&F|CIP)[\s\-–\\]*(JPY|USD|EUR|CNY|GBP|THB|AUD)[\s\-–\\]*" + AMT, flat)
    if m:
        out["inv_cur"], out["inv_amt"] = m.group(1), to_num(m.group(2))
    m = re.search(r"数量\s*(?:[\(（]\s*[12１２]\s*[\)）])?\s*" + AMT + r"\s*(KGM|KG)\b", flat) or \
        re.search(r"(?<![\d.,])" + AMT + r"\s*(KGM|KG)\b", flat)
    if m:
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


def read_docs(pdf: Path, suppliers=(), products=None) -> tuple[list[dict], list[dict]]:
    """1つのPDFから、許可通知書と送金計算書を全ページ分読む。"""
    products = products if products is not None else DEFAULT_PRODUCTS
    permits: dict[str, dict] = {}
    remits: list[dict] = []
    with pdfplumber.open(pdf) as doc:
        texts = [(p.extract_text(layout=True) or "") for p in doc.pages]
    for i, t in enumerate(texts):
        items = None
        if not t.strip():
            if i >= OCR_MAX_PAGES:
                continue
            items = ocr_items(pdf, i)
            t = "\n".join(x[2] for x in items)
        if is_remit_text(t, pdf.name):
            r = parse_remit(t, suppliers)
            if r and not any((r["date"], r["cur"], r["amount"]) == (x["date"], x["cur"], x["amount"]) for x in remits):
                remits.append(r)
            continue
        info = parse_items(items, suppliers) if items is not None else parse_page_text(t, suppliers)
        if info is None:
            continue
        info.update(extract_extras(t, products))
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
    if not out and not remits:
        raise ValueError("許可通知書・送金計算書が見つかりません")
    return out, remits


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


def build_groups(permits: list[dict], remits: list[dict], window: int = 150) -> list[dict]:
    """送金ごとに、金額の合う許可通知書を結びつける。結びつかないものも残す。"""
    free = list(permits)
    groups = []
    for rem in sorted(remits, key=lambda r: r["date"]):
        cands = []
        for p in free:
            if p.get("inv_amt") is None or p.get("inv_cur") != rem["cur"]:
                continue
            if p["permit"] and abs((p["permit"] - rem["date"]).days) > window:
                continue
            if (rem["supplier_sure"] and p["shipper_sure"] and p["shipper"]
                    and norm_name(p["shipper"]) != norm_name(rem["supplier"])):
                continue
            cands.append(p)
        cands.sort(key=lambda p: abs((p["permit"] - rem["date"]).days) if p["permit"] else 9999)
        sols = find_subsets(cands, rem["amount"])
        chosen = sols[0] if sols else []
        ambiguous = len({tuple(sorted(id(x) for x in s)) for s in sols}) > 1
        for p in chosen:
            free.remove(p)
        groups.append({"rem": rem, "items": chosen, "ambiguous": ambiguous})
    groups.extend({"rem": None, "items": [p], "ambiguous": False} for p in free)

    def key(g):
        ds = [g["rem"]["date"]] if g["rem"] else [p["permit"] for p in g["items"] if p["permit"]]
        return min(ds) if ds else datetime.max
    return sorted(groups, key=key)


def describe_groups(groups: list[dict]) -> None:
    for g in groups:
        rem, items = g["rem"], g["items"]
        if rem:
            amt = f"{rem['amount']:,.2f}" if rem["cur"] != "JPY" else f"{int(rem['amount']):,}"
            print(f"■ 送金 {rem['date']:%Y/%m/%d}  {rem['cur']} {amt}  {rem['supplier'] or '(受取人不明)'}"
                  + (f"  → 許可通知書 {len(items)} 件" if items else "  → 金額の合う許可通知書が見つかりません(要確認)")
                  + ("  ※合う組み合わせが複数あります(要確認)" if g["ambiguous"] else ""))
        else:
            print("■ 送金計算書が見つかりません(送金の欄は空欄・黄色になります)")
        for p in sorted(items, key=lambda x: x['permit'] or datetime.max):
            d = f"{p['permit']:%Y/%m/%d}" if p["permit"] else "(許可日なし)"
            a_ = f"{p['inv_cur']} {p['inv_amt']:,.2f}".rstrip("0").rstrip(".") if p.get("inv_amt") is not None else "金額不明"
            flag = "" if p["shipper_sure"] else "  ※仕出人要確認"
            print(f"    申告番号 {p['decl']}  許可日 {d}  {a_}  数量 {p['qty'] or '不明'}  {p['shipper']}{flag}")
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
    wb = openpyxl.load_workbook(TPL_LEDGER)
    for ws in wb:
        # 見本データを消し、見出しだけ残す(「例」行は送金ありのみ残す)
        first = 5 if ws.title == SHEET_SEND else 4
        for row in ws.iter_rows(min_row=first, max_row=ws.max_row):
            for c in row:
                c.value = None
    wb[SHEET_SEND]["B1"] = period_label(period)
    wb[SHEET_NOSEND]["A1"] = period_label(period)
    wb.save(path)


def next_no(ws, col: int, prefix: str, period: int) -> str:
    pat = re.compile(rf"^{re.escape(prefix)}{period}-(\d+)$")
    nums = [int(m.group(1)) for r in range(5 if ws.title == SHEET_SEND else 4, ws.max_row + 1)
            if (v := ws.cell(r, col).value) and (m := pat.match(str(v)))]
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
def make_cover(out: Path, no: str, info: dict, nosend: str | None, remit_yen=None, many=False):
    wb = openpyxl.load_workbook(TPL_COVER)
    if nosend is None:
        wb.remove(wb.worksheets[1])
        ws = wb.worksheets[0]
        ws["E16"] = None  # 見本の「乙仲 No.」を消す
        ymd = ("E10", "I10", "K10")
        if remit_yen:
            ws["J48"] = int(remit_yen)
            ws["J48"].number_format = "#,##0"
    else:
        wb.remove(wb.worksheets[0])
        ws = wb.worksheets[0]
        ymd = ("E10", "H10", "J10")
        if nosend in ("着払", "無償"):
            ws["H1" if nosend == "着払" else "J1"].fill = HILITE
    ws["E3"] = no
    ws["E5"] = info["decl"] if many else info["decl"].replace(" ", "")
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


def run_all(a, root: Path, paths, confirm=None) -> int:
    """PDFを読み、送金ごとにまとめて表示し、(確認のうえ)台帳と表紙に保存する。保存した件数を返す。"""
    suppliers, products = load_suppliers(root), load_products(root)
    permits, remits, skipped, seen = [], [], 0, set()
    for pdf in collect(paths):
        try:
            ps, rs = read_docs(pdf, suppliers, products)
        except Exception:
            skipped += 1  # 請求書・到着案内など、対象外のPDF
            continue
        for p in ps:
            if p["decl"] not in seen:
                seen.add(p["decl"])
                p["pdf"] = pdf
                permits.append(p)
        for r in rs:
            r["pdf"] = pdf
            remits.append(r)
    print(f"許可通知書 {len(permits)} 件、送金計算書 {len(remits)} 件を読み取りました"
          f"(それ以外のPDF {skipped} 件はスキップ)\n")
    if not permits and not remits:
        return 0
    if a.nosend:
        groups = [{"rem": None, "items": [p], "ambiguous": False}
                  for p in sorted(permits, key=lambda x: x["permit"] or datetime.max)]
    else:
        groups = build_groups(permits, remits)
    describe_groups(groups)
    if a.dry_run:
        return len(groups)
    if confirm and not confirm(len(groups)):
        return 0
    return sum(save_group(a, root, g) for g in groups)


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
    n = run_all(a, out, [src], confirm)
    if not n:
        messagebox.showinfo("結果", "保存したものはありません")
        return 1
    messagebox.showinfo("完了", f"{n} 件を保存しました。\n\n{out}")
    try:
        os.startfile(out)
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
    cover = folder / "表紙" / f"表紙_{no}.xlsx"
    if not cover.exists():
        make_cover(cover, no, info, a.nosend)
    elif shared:  # 同月まとめの2件目以降: 既存の表紙の申告番号欄に追記する
        cwb = openpyxl.load_workbook(cover)
        cws = cwb.worksheets[0]
        cur = str(cws["E5"].value or "")
        if want not in cur:
            cws["E5"] = f"{cur}、{want}" if cur else want
            cwb.save(cover)
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
    ws = wb[SHEET_SEND]

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
        no = shared or next_no(ws, 3, "", period)
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
            put(row, 5, p["product"])
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
        warn = [c for c in (4, 5, 6, 7, 9, 10, 12, 13, 14, 15) if ws.cell(row, c).value in (None, "")]
        if p and not p["shipper_sure"]:
            warn += [4, 14]
        if rem and not p and not rem["supplier_sure"]:
            warn += [4]
        if ambiguous:
            warn += [3]
        for c in warn:
            ws.cell(row, c).fill = HILITE
    wb.save(ledger)

    if fallback:
        fmap[fkey] = no
        fmap_path.write_text(json.dumps(fmap, ensure_ascii=False, indent=1), encoding="utf8")

    # 表紙(管理番号ごとに1枚。申告番号は全件を並べる)
    base = dict(items[0]) if items else {"decl": "", "permit": rem["date"], "shipper": rem["supplier"]}
    base["decl"] = "、".join(p["decl"].replace(" ", "") for p in items) if items else ""
    dated = [p["permit"] for p in items if p["permit"]]
    if dated:
        base["permit"] = min(dated)
    cover = folder / "表紙" / f"表紙_{no}.xlsx"
    make_cover(cover, no, base, None, rem["yen"] if rem else None, many=True)
    d = rem["date"].strftime("%Y/%m/%d") if rem else "送金なし"
    print(f"[OK] {no}  送金 {d}  許可通知書 {len(items)} 件" + ("  ※黄色の欄を確認" if (not rem or ambiguous) else ""))
    return 1



if __name__ == "__main__":
    sys.exit(main())
