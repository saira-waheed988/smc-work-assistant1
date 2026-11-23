"""Area Tracker - comparison logic (no Streamlit code here, so it is easy to test)."""
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

GREEN = PatternFill("solid", start_color="C6EFCE")   # New
RED = PatternFill("solid", start_color="FFC7CE")     # Removed
YELLOW = PatternFill("solid", start_color="FFEB9C")  # Modified
ORANGE = PatternFill("solid", start_color="F4B183")  # Modified cell itself
HEADER = PatternFill("solid", start_color="1F3864")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FONT = "Arial"

STATUS_FILL = {"New": GREEN, "Removed": RED, "Modified": YELLOW}


def list_sheets(file) -> list[str]:
    file.seek(0)
    return load_workbook(file, read_only=True).sheetnames


def read_sheet(file, sheet: str, header_row: int = 1) -> pd.DataFrame:
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet, header=header_row - 1, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df.reset_index(drop=True)


def _norm_key(v) -> str:
    return "" if pd.isna(v) else str(v).strip().upper()


def _is_blank(v) -> bool:
    return v is None or (not isinstance(v, str) and pd.isna(v)) or (isinstance(v, str) and v.strip() == "")


def _same(a, b, tol: float) -> bool:
    if _is_blank(a) and _is_blank(b):
        return True
    if _is_blank(a) or _is_blank(b):
        return False
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip().upper() == str(b).strip().upper()


def compare(old: pd.DataFrame, new: pd.DataFrame, key: str, tol: float = 0.0,
            ignore_cols: list[str] | None = None) -> dict:
    """Compare two area lists by a key column. Returns dict of DataFrames."""
    ignore = set(ignore_cols or [])
    compare_cols = [c for c in new.columns if c in old.columns and c != key and c not in ignore]

    notes = []
    for label, df in (("Previous", old), ("Updated", new)):
        keys = df[key].map(_norm_key)
        if (keys == "").any():
            notes.append(f"{label}: {(keys == '').sum()} row(s) have empty '{key}' - skipped.")
        dups = sorted(set(keys[keys.duplicated(keep=False) & (keys != "")]))
        if dups:
            notes.append(f"{label}: duplicate '{key}' values (first one used): {', '.join(dups[:15])}"
                         + (" ..." if len(dups) > 15 else ""))
    only_old = [c for c in old.columns if c not in new.columns]
    only_new = [c for c in new.columns if c not in old.columns]
    if only_old:
        notes.append(f"Columns only in Previous (not compared): {', '.join(only_old)}")
    if only_new:
        notes.append(f"Columns only in Updated (not compared): {', '.join(only_new)}")

    def index(df):
        d = df.copy()
        d["__k"] = d[key].map(_norm_key)
        d = d[d["__k"] != ""].drop_duplicates("__k", keep="first")
        return d.set_index("__k")

    o, n = index(old), index(new)
    rows, changes = [], []

    for k in n.index:
        nr = n.loc[k]
        if k not in o.index:
            rows.append({**{c: nr.get(c) for c in new.columns}, "Status": "New", "Changed Columns": ""})
            continue
        orow = o.loc[k]
        diffs = [c for c in compare_cols if not _same(orow[c], nr[c], tol)]
        for c in diffs:
            ov, nv = orow[c], nr[c]
            delta = None
            try:
                delta = float(nv) - float(ov)
            except (TypeError, ValueError):
                pass
            changes.append({key: nr[key], "Column": c, "Previous": ov, "Updated": nv, "Difference": delta})
        rows.append({**{c: nr.get(c) for c in new.columns},
                     "Status": "Modified" if diffs else "Unchanged",
                     "Changed Columns": ", ".join(diffs)})

    for k in o.index:
        if k not in n.index:
            orow = o.loc[k]
            rows.append({**{c: orow.get(c) for c in old.columns}, "Status": "Removed", "Changed Columns": ""})

    result = pd.DataFrame(rows)
    cols = [c for c in result.columns if c not in ("Status", "Changed Columns")]
    result = result[["Status"] + cols + ["Changed Columns"]]
    order = {"New": 0, "Modified": 1, "Removed": 2, "Unchanged": 3}
    result = result.sort_values("Status", key=lambda s: s.map(order), kind="stable").reset_index(drop=True)

    return {"result": result, "changes": pd.DataFrame(changes), "notes": notes, "compare_cols": compare_cols}


def summary_counts(result: pd.DataFrame) -> dict:
    vc = result["Status"].value_counts()
    return {s: int(vc.get(s, 0)) for s in ("New", "Modified", "Removed", "Unchanged")}


def totals(result: pd.DataFrame, area_col: str | None) -> dict | None:
    """Total area before/after if an area column is chosen."""
    if not area_col or area_col not in result.columns:
        return None
    vals = pd.to_numeric(result[area_col], errors="coerce").fillna(0)
    st = result["Status"]
    new_total = vals[st.isin(["New", "Modified", "Unchanged"])].sum()
    return {"added": vals[st == "New"].sum(), "removed": vals[st == "Removed"].sum(), "updated_total": new_total}


def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _autowidth(ws, minw=10, maxw=45):
    for col in ws.columns:
        length = max((len(str(c.value)) for c in col if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(minw, min(maxw, length + 3))


def to_excel(out: dict, key: str, area_col: str | None, old_name: str, new_name: str) -> bytes:
    result, changes = out["result"], out["changes"]
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        result.to_excel(xw, sheet_name="Comparison", index=False)
        (changes if len(changes) else pd.DataFrame(columns=[key, "Column", "Previous", "Updated", "Difference"])
         ).to_excel(xw, sheet_name="Changes Detail", index=False)
        pd.DataFrame().to_excel(xw, sheet_name="Summary", index=False)
    buf.seek(0)
    wb = load_workbook(buf)

    # --- Comparison sheet: coloured rows, orange cells for the changed values
    ws = wb["Comparison"]
    _style_header(ws, 1, ws.max_column)
    headers = [c.value for c in ws[1]]
    for r in range(2, ws.max_row + 1):
        status = ws.cell(r, 1).value
        changed = [x.strip() for x in str(ws.cell(r, ws.max_column).value or "").split(",") if x.strip()]
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name=FONT, size=10, bold=(c == 1))
            cell.border = BORDER
            if status in STATUS_FILL:
                cell.fill = STATUS_FILL[status]
            if status == "Modified" and headers[c - 1] in changed:
                cell.fill = ORANGE
                cell.font = Font(name=FONT, size=10, bold=True)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    _autowidth(ws)

    # --- Changes detail
    wc = wb["Changes Detail"]
    _style_header(wc, 1, wc.max_column)
    for row in wc.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name=FONT, size=10)
            cell.border = BORDER
    wc.freeze_panes = "A2"
    _autowidth(wc)

    # --- Summary sheet (live formulas on the Comparison sheet)
    sm = wb["Summary"]
    wb.move_sheet("Summary", offset=-2)
    last = ws.max_row
    sm["A1"] = "Area Comparison Summary"
    sm["A1"].font = Font(name=FONT, size=14, bold=True)
    sm["A3"], sm["B3"] = "Previous file", old_name
    sm["A4"], sm["B4"] = "Updated file", new_name
    sm["A5"], sm["B5"] = "Key column", key
    for r in (3, 4, 5):
        sm.cell(r, 1).font = Font(name=FONT, bold=True)
        sm.cell(r, 2).font = Font(name=FONT)
    sm["A7"], sm["B7"], sm["C7"] = "Status", "Count", "Colour"
    _style_header(sm, 7, 3)
    for i, (s, fill) in enumerate([("New", GREEN), ("Modified", YELLOW), ("Removed", RED), ("Unchanged", None)]):
        r = 8 + i
        sm.cell(r, 1, s)
        sm.cell(r, 2, f"=COUNTIF(Comparison!$A$2:$A${last},A{r})")
        sm.cell(r, 3, {"New": "Green", "Modified": "Yellow", "Removed": "Red", "Unchanged": "No highlight"}[s])
        for c in (1, 2, 3):
            sm.cell(r, c).font = Font(name=FONT)
            sm.cell(r, c).border = BORDER
            if fill:
                sm.cell(r, c).fill = fill
    sm["A12"], sm["B12"] = "Total rows", "=SUM(B8:B11)"
    sm["A12"].font = sm["B12"].font = Font(name=FONT, bold=True)

    if area_col and area_col in headers:
        col = get_column_letter(headers.index(area_col) + 1)
        rng = f"Comparison!${col}$2:${col}${last}"
        stat = f"Comparison!$A$2:$A${last}"
        sm["A14"], sm["B14"] = f"Area totals ({area_col})", "Value"
        _style_header(sm, 14, 2)
        items = [
            ("Added (New areas)", f'=SUMIFS({rng},{stat},"New")'),
            ("Removed areas", f'=SUMIFS({rng},{stat},"Removed")'),
            ("Updated list total", f'=SUMIFS({rng},{stat},"New")+SUMIFS({rng},{stat},"Modified")'
                                   f'+SUMIFS({rng},{stat},"Unchanged")'),
        ]
        for i, (lbl, f) in enumerate(items):
            sm.cell(15 + i, 1, lbl).font = Font(name=FONT)
            c = sm.cell(15 + i, 2, f)
            c.font = Font(name=FONT)
            c.number_format = "#,##0.00"
            sm.cell(15 + i, 1).border = c.border = BORDER

    r0 = 20
    sm.cell(r0, 1, "Notes / warnings").font = Font(name=FONT, bold=True)
    for i, n in enumerate(out["notes"] or ["None"]):
        sm.cell(r0 + 1 + i, 1, n).font = Font(name=FONT)
    sm.cell(r0 + len(out["notes"] or [1]) + 3, 1,
            "Tool flags differences only - final verification by the responsible engineer.").font = \
        Font(name=FONT, italic=True, color="7F7F7F")
    sm.column_dimensions["A"].width = 30
    sm.column_dimensions["B"].width = 38
    sm.column_dimensions["C"].width = 16

    wb.active = 0
    out_buf = BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()
