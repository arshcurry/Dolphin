import os
import re
from collections import defaultdict
from typing import List, Tuple

# ====== Folders ======
ROOT = os.getcwd()
ALL_REPORTS_DIR = os.path.join(ROOT, "All_reports")
OUT_DIR = os.path.abspath(os.path.join(ALL_REPORTS_DIR, "Consolidated"))
os.makedirs(OUT_DIR, exist_ok=True)

# Filter to one month like "08-2025", or None for all
TARGET_MONTH_YEAR = None

# ====== Keys & labels ======
KEYS = [
    "BC_PTD", "ARR_I", "ARR_E", "MS12", "TB1", "TB", "BS", "IS", "AR", "PR", "GL", "L", "BC"
]
LABELS = {
    "BC":     "Budget Comparison",
    "BC_PTD": "Budget Comparison (with PTD)",
    "TB1":    "Trial Balance",
    "TB":     "Trial Balance",
    "BS":     "Balance Sheet",
    "IS":     "Income Statement",
    "ARR_I":  "Affordable Receivable Aging Summary",
    "ARR_E":  "Affordable Receivable Aging Summary (Exclude Affordable)",
    "AR":     "Affordable Rent Roll with Lease Charges",
    "PR":     "Rent Roll with Lease Charges",
    "GL":     "General Ledger",
    "L":      "Legal",
    "MS12":   "12 month Statement",
}

SEQUENCE_SINGLE   = ["BC", "TB1", "TB", "BS", "IS", "ARR_I", "ARR_E", "AR", "GL", "L"]
SEQUENCE_NUMBERED = ["BC", "TB1", "TB", "BS", "IS", "ARR_I", "ARR_E", "PR", "GL", "MS12", "L"]
SEQUENCE_MULTI_STATIC_HEAD = ["BC_PTD", "TB1", "TB", "BS", "IS", "ARR_I", "ARR_E"]
SEQUENCE_MULTI_STATIC_TAIL = ["GL", "L"]

# ====== Filename parsing ======
FILENAME_RE = re.compile(
    r"^(?P<code>.+?)_(?P<date>\d{2}-(?:\d{2}-)?\d{4})_(?P<suffix>[A-Z0-9]+(?:_[A-Z0-9]+)?)(?P<dup>\d+)?$",
    re.IGNORECASE
)

def extract_month_year(date_str: str) -> str:
    parts = date_str.split("-")
    return date_str if len(parts) == 2 else f"{parts[0]}-{parts[2]}"

def month_dot(month_year: str) -> str:
    return month_year.replace("-", ".")

def detect_key_from_suffix(suffix: str) -> str:
    s = suffix.upper()
    for key in sorted(KEYS, key=len, reverse=True):
        if s == key or re.fullmatch(key + r"\d+", s):
            return key
    if s == "ARR":
        return "AR"
    return None

def is_numbered_property(code: str) -> bool:
    return bool(re.search(r"\d+$", code)) and "^" not in code

def is_multi_property(code: str) -> bool:
    return "^" in code

def looks_like_our_output(name: str) -> bool:
    return ("_Mgmt Report_" in name) or ("_CONSOLIDATED" in name)

# ====== Scan All_reports ======
def scan_folder(folder: str):
    rows = []
    for name in os.listdir(folder):
        if not name.lower().endswith(".xlsx"):
            continue
        if looks_like_our_output(name):
            continue
        base = os.path.splitext(name)[0]
        m = FILENAME_RE.match(base)
        if not m:
            continue
        code   = m.group("code")
        date_s = m.group("date")
        suffix = (m.group("suffix") or "").upper()
        key    = detect_key_from_suffix(suffix)
        if key is None:
            continue
        path = os.path.join(folder, name)
        rows.append({
            "code": code,
            "date": date_s,
            "month_year": extract_month_year(date_s),
            "key": key,
            "suffix": suffix,
            "path": path,
            "mtime": os.path.getmtime(path),
            "name": name,
        })
    return rows

# ====== Excel COM helpers (preserve formatting 1:1) ======
def ensure_unique_sheet_name(xl_wb, base: str) -> str:
    name = base[:31]
    existing = {ws.Name for ws in xl_wb.Worksheets}
    if name not in existing:
        return name
    i = 2
    while True:
        cand = (f"{name[:29]} {i}") if len(name) > 29 else f"{name} {i}"
        if cand not in existing:
            return cand
        i += 1

def copy_first_sheet_via_excel(excel, src_path: str, dest_wb, new_name: str):
    src_wb = excel.Workbooks.Open(src_path, ReadOnly=True, UpdateLinks=0)
    try:
        src_ws = src_wb.Worksheets(1)
        after_ws = dest_wb.Worksheets(dest_wb.Worksheets.Count)
        src_ws.Copy(After=after_ws)  # creates a new sheet inside dest_wb
        new_ws = dest_wb.Worksheets(dest_wb.Worksheets.Count)
        new_ws.Name = ensure_unique_sheet_name(dest_wb, new_name)
        return new_ws
    finally:
        src_wb.Close(SaveChanges=False)

def sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name)

def unique_path(folder: str, filename: str) -> str:
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(folder, filename)
    i = 1
    while os.path.exists(candidate):
        candidate = os.path.join(folder, f"{base}({i}){ext}")
        i += 1
    return candidate

# --- helpers for header repair ---
def last_used_col(ws, header_row_guess=5):
    xlToLeft = -4159
    try:
        return ws.Cells(header_row_guess, ws.Columns.Count).End(xlToLeft).Column
    except Exception:
        # fallback: scan first 100 cols
        last = 1
        for c in range(1, 101):
            if str(ws.Cells(header_row_guess, c).Value).strip():
                last = c
        return last

def extend_top_merges(ws, last_col, rows=(1, 2, 3)):
    """Re-span the top heading merges to the current last_col."""
    xlCenter = -4108
    for r in rows:
        # find first non-empty cell on the row
        start_c = None
        for c in range(1, min(60, last_col) + 1):
            v = ws.Cells(r, c).Value
            if v is not None and str(v).strip() != "":
                start_c = c
                break
        if not start_c:
            continue

        cell = ws.Cells(r, start_c)
        # get original left edge of merge (if any)
        if cell.MergeCells:
            area = cell.MergeArea
            left_col = area.Column
            try:
                area.UnMerge()
            except Exception:
                pass
        else:
            left_col = start_c

        rng = ws.Range(ws.Cells(r, left_col), ws.Cells(r, last_col))
        try:
            rng.Merge()
            rng.HorizontalAlignment = xlCenter
        except Exception:
            # if some conflicting merge exists, skip quietly
            pass

# --------- ONLY tweak the header for Budget Comparison (not PTD) ----------
def add_mtd_and_fix_header(ws, excel):
    """
    • Find the header row with 'Annual' and 'Note/Notes'
    • Insert a column at Notes -> label it 'MTD'
    • Rename the (shifted) Notes to 'YTD'
    • Re-extend the 3-line title merges so the banner looks identical
    (No cell values are copied anywhere.)
    """
    try:
        # 1) find header row and 'Notes' column
        header_row = None
        note_col = None
        for r in range(1, 21):
            vals = []
            for c in range(1, 80):
                v = ws.Cells(r, c).Value
                vals.append(str(v).strip() if v is not None else "")
            if "Annual" in vals and any(v.lower() in ("note", "notes") for v in vals):
                header_row = r
                idx = next(i for i, v in enumerate(vals) if v.lower() in ("note", "notes"))
                note_col = idx + 1
                break
        if not (header_row and note_col):
            return  # nothing to do

        # 2) insert one column at current Notes col (inserts LEFT)
        ws.Columns(note_col).Insert()

        # 3) label headers (no data copy)
        ws.Cells(header_row, note_col).Value = "MTD"
        ws.Cells(header_row, note_col + 1).Value = "YTD"   # old Notes shifted right

        # match width only (no format or data copy)
        try:
            ws.Columns(note_col).ColumnWidth = ws.Columns(note_col + 1).ColumnWidth
        except Exception:
            pass

        # 4) re-extend the three top merged banners so they span to new last col
        last_col = last_used_col(ws, header_row_guess=header_row)
        extend_top_merges(ws, last_col, rows=(1, 2, 3))

    except Exception:
        # never fail consolidation due to a cosmetic header change
        pass

# ====== Main consolidation ======
def consolidate():
    records = scan_folder(ALL_REPORTS_DIR)
    if not records:
        print(f"No .xlsx files found in {ALL_REPORTS_DIR}")
        return

    by_code_month = defaultdict(list)
    by_code_month_key = defaultdict(list)
    for r in records:
        by_code_month[(r["code"], r["month_year"])].append(r)
        by_code_month_key[(r["code"], r["month_year"], r["key"])].append(r)

    targets = sorted(set((r["code"], r["month_year"]) for r in records))
    if TARGET_MONTH_YEAR:
        targets = [t for t in targets if t[1] == TARGET_MONTH_YEAR]

    # Skip single-property outputs when that code is part of a multi for that month
    multi_subcodes_by_month = defaultdict(set)
    for code, month in targets:
        if is_multi_property(code):
            for sub in code.split("^"):
                multi_subcodes_by_month[month].add(sub)

    filtered_targets: List[Tuple[str, str]] = []
    for code, month in targets:
        if not is_multi_property(code) and not is_numbered_property(code):
            if code in multi_subcodes_by_month[month]:
                continue
        filtered_targets.append((code, month))

    import win32com.client as win32
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.DefaultFilePath = OUT_DIR

    try:
        for code, month_year in filtered_targets:
            category = "multi" if is_multi_property(code) else ("numbered" if is_numbered_property(code) else "single")
            print(f"\n🔧 Consolidating: {code}  |  {month_year}  |  category={category}")

            order = []
            if category == "single":
                for k in SEQUENCE_SINGLE:
                    order.append((k, code, None))
            elif category == "numbered":
                for k in SEQUENCE_NUMBERED:
                    order.append((k, code, None))
            else:
                subs = code.split("^")
                for k in SEQUENCE_MULTI_STATIC_HEAD:
                    order.append((k, code, None))
                for sub in subs:
                    order.append(("AR", sub, f" ({sub})"))
                for sub in subs:
                    order.append(("BC", sub, f" ({sub})"))
                for k in SEQUENCE_MULTI_STATIC_TAIL:
                    order.append((k, code, None))

            dest_wb = excel.Workbooks.Add()
            initial_sheet_names = [ws.Name for ws in dest_wb.Worksheets]
            any_copied = False
            sheet_index = 1

            for key, lookup_code, suffix_note in order:
                if key == "PR":
                    cands = [r for r in by_code_month.get((lookup_code, month_year), []) if r["key"] == "PR"]
                else:
                    cands = by_code_month_key.get((lookup_code, month_year, key), [])

                if not cands:
                    continue

                best = max(cands, key=lambda r: r["mtime"])
                label = LABELS.get(key, key)
                sheet_title = f"{sheet_index:02d} {label}{suffix_note or ''}"

                try:
                    new_ws = copy_first_sheet_via_excel(excel, best["path"], dest_wb, sheet_title)
                    any_copied = True
                    sheet_index += 1

                    # --- Only for regular Budget Comparison (NOT PTD) ---
                    if key == "BC" and new_ws is not None:
                        add_mtd_and_fix_header(new_ws, excel)

                except Exception as e:
                    print(f"   ⚠️ Failed to copy {best['name']}: {e}")

            if any_copied:
                try:
                    for nm in initial_sheet_names:
                        for ws in list(dest_wb.Worksheets):
                            if ws.Name == nm:
                                try: ws.Delete()
                                except Exception: pass
                except Exception:
                    pass
            else:
                print("   ⚠️ No files found for this code/month. Skipping output.")
                dest_wb.Close(SaveChanges=False)
                continue

            out_name = sanitize_filename(f"{code}_Mgmt Report_{month_dot(month_year)}_Sent.xlsx")
            out_path = unique_path(OUT_DIR, out_name)
            os.makedirs(os.path.dirname(out_path), exist_ok=True)

            try:
                dest_wb.SaveCopyAs(out_path)
                if os.path.exists(out_path):
                    print(f"✅ Saved to: {out_path}")
                else:
                    print("⛔ SaveCopyAs returned but file not found at expected path.")
            except Exception as e:
                print(f"⛔ SaveCopyAs failed: {e}")

            dest_wb.Close(SaveChanges=False)

    finally:
        excel.DisplayAlerts = True
        excel.Quit()

if __name__ == "__main__":
    consolidate()