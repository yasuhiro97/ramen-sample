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
SHEET_SEND = "海外送金あり"
SHEET_NOSEND = "海外送金なし（乙仲・無償・着払）"
HILITE = PatternFill("solid", fgColor="FFFF00")


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


def load_suppliers(root: Path) -> list[str]:
    """仕入先名の一覧: 台帳(全期)の仕入先名 + 仕入先名.txt(1行1社)。"""
    names = set()
    f = root / "仕入先名.txt"
    if f.exists():
        names |= {l.strip() for l in f.read_text(encoding="utf8").splitlines() if l.strip()}
    for lg in root.glob("第*期/輸入明細一覧_*.xlsx"):
        wb = openpyxl.load_workbook(lg, read_only=True)
        for ws, col in ((wb[SHEET_SEND], 4), (wb[SHEET_NOSEND], 3)):
            for r in ws.iter_rows(min_row=4, min_col=col, max_col=col, values_only=True):
                if r[0]:
                    names.add(str(r[0]).strip())
    return sorted(names)


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
def make_cover(out: Path, no: str, info: dict, nosend: str | None):
    wb = openpyxl.load_workbook(TPL_COVER)
    if nosend is None:
        wb.remove(wb.worksheets[1])
        ws = wb.worksheets[0]
        ws["E16"] = None  # 見本の「乙仲 No.」を消す
        ymd = ("E10", "I10", "K10")
    else:
        wb.remove(wb.worksheets[0])
        ws = wb.worksheets[0]
        ymd = ("E10", "H10", "J10")
        if nosend in ("着払", "無償"):
            ws["H1" if nosend == "着払" else "J1"].fill = HILITE
    ws["E3"] = no
    ws["E5"] = info["decl"].replace(" ", "")
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


def gui() -> int:
    """引数なしで起動したとき(ダブルクリック用)。画面でフォルダ・事業所名・種別を選ぶ。"""
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
    a = Namespace(office=office, nosend=nosend, period=None, dry_run=True, out=out)
    suppliers = load_suppliers(out)
    print(f"読み取り中: {src}\n(スキャンPDFが多いと時間がかかります)\n")
    found = []
    for pdf in collect([src]):
        try:
            for info in read_permits(pdf, suppliers):
                found.append((pdf, info))
        except Exception as e:
            print(f"[対象外] {pdf.name}: {e}")
    if not found:
        messagebox.showinfo("結果", "許可通知書が見つかりませんでした")
        return 1
    print()
    for pdf, info in found:
        process(a, out, pdf, info)  # dry_run=True: 内容を表示するだけ
    if not messagebox.askyesno("保存の確認", f"{len(found)} 件の許可通知書を読み取りました。\n"
                               "画面の内容を確認して、台帳と表紙を作成しますか?"):
        return 0
    a.dry_run = False
    n = sum(process(a, out, pdf, info) for pdf, info in found)
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
    ok = 0
    suppliers = load_suppliers(root)
    for pdf in collect(a.pdf):
        try:
            infos = read_permits(pdf, suppliers)
        except Exception as e:
            print(f"[NG] {pdf.name}: {e}", file=sys.stderr)
            continue
        for info in infos:
            ok += process(a, root, pdf, info)
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
    if a.nosend:
        no = next_no(ws, 2, "経", period)
        row = add_nosend(ws, no, a.office, info)
    else:
        no = next_no(ws, 3, "", period)
        row = add_send(ws, no, a.office, info)
    if not info["shipper_sure"]:
        for c in ((row, 3), (row, 10)) if a.nosend else ((row, 4), (row, 14)):
            ws.cell(*c).fill = HILITE
        print(f"      仕出人『{info['shipper']}』は照合できませんでした(黄色の欄を要確認)")
    wb.save(ledger)
    make_cover(folder / "表紙" / f"表紙_{no}.xlsx", no, info, a.nosend)
    d = f"{info['permit']:%Y/%m/%d}" if info["permit"] else "(許可日は手入力)"
    print(f"[OK] {pdf.name} -> {no}  申告番号 {info['decl']}  許可日 {d}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
