"""Drawing Manager - drawing register analysis (no Streamlit code here)."""
import re
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from area_logic import BORDER, FONT, GREEN, ORANGE, RED, YELLOW, _autowidth, _style_header

GREY = PatternFill("solid", start_color="E7E6E6")
STD = ["Drawing No", "Title", "Floor", "Category", "Revision", "Date"]


# ---------------------------------------------------------------- helpers
def _blank(v) -> bool:
    return v is None or (not isinstance(v, str) and pd.isna(v)) or str(v).strip() == ""


def _norm(v) -> str:
    return "" if _blank(v) else re.sub(r"\s+", " ", str(v).strip()).upper()


def _txt(v) -> str:
    return "" if _blank(v) else str(v).strip()


def rev_rank(rev):
    """Revision -> sortable number. A=1,B=2..Z=26,AA=27 ; 0,1,2 ; 'REV C', 'R02', 'P01' all understood."""
    if _blank(rev):
        return None
    s = str(rev).strip().upper()
    if s in ("-", "NA", "N/A", "NIL"):
        return None
    s = re.sub(r"^REV(ISION)?[\s\.\-_]*", "", s)
    s = re.sub(r"^[RPC](?=\d)", "", s)
    if s.isdigit():
        return int(s)
    if s.isalpha() and len(s) <= 3:
        v = 0
        for ch in s:
            v = v * 26 + (ord(ch) - 64)
        return v
    m = re.search(r"\d+", s)
    return int(m.group()) if m else None


def read_register(file, sheet: str, header_row: int = 1) -> pd.DataFrame:
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet, header=header_row - 1, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    df["_row"] = df.index + header_row + 1          # real Excel row number
    data_cols = [c for c in df.columns if c != "_row"]
    df = df.dropna(how="all", subset=data_cols)
    return df.reset_index(drop=True)


def standardize(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)
    for col in STD:
        src = mapping.get(col)
        out[col] = df[src] if src and src in df.columns else None
    if not mapping.get("Category"):
        out["Category"] = "All drawings"
    out["_row"] = df["_row"]
    return out


def found_values(reg: pd.DataFrame, col: str) -> list[str]:
    """Unique values (original spelling) in order of appearance, ignoring case/space differences."""
    seen, res = set(), []
    for v in reg[col]:
        n = _norm(v)
        if n and n not in seen:
            seen.add(n)
            res.append(_txt(v))
    return res


# --------------------------------------------------------------- analysis
def analyze(reg: pd.DataFrame, expected_floors: list[str], required_cats: list[str]) -> dict:
    d = reg.copy()
    d["_key"] = d["Drawing No"].map(_norm)
    d["_floor"] = d["Floor"].map(_norm)
    d["_cat"] = d["Category"].map(_norm)
    d["_rank"] = pd.to_numeric(d["Revision"].map(rev_rank), errors="coerce")
    try:
        d["_date"] = pd.to_datetime(d["Date"], errors="coerce", format="mixed")
    except (TypeError, ValueError):
        d["_date"] = pd.to_datetime(d["Date"], errors="coerce")
    d["_revkey"] = [f"#{int(r)}" if pd.notna(r) else "T:" + _norm(t) for r, t in zip(d["_rank"], d["Revision"])]

    issues, flags = [], {}

    def add(sev, issue, row, detail, tag):
        issues.append({"Severity": sev, "Issue": issue, "Drawing No": _txt(row["Drawing No"]),
                       "Floor": _txt(row["Floor"]), "Register Row": int(row["_row"]), "Details": detail})
        flags.setdefault(int(row["_row"]), [])
        if tag and tag not in flags[int(row["_row"])]:
            flags[int(row["_row"])].append(tag)

    valid = d[d["_key"] != ""]

    for _, r in d[d["_key"] == ""].iterrows():
        add("Error", "Missing drawing number", r, f"Row has title '{_txt(r['Title'])}' but no drawing number.", "No drawing no")

    for _, r in valid[valid.duplicated(["_key", "_revkey"], keep=False)].iterrows():
        add("Error", "Duplicate drawing no + revision", r,
            f"{_txt(r['Drawing No'])} Rev {_txt(r['Revision']) or '-'} is listed more than once.", "Duplicate")

    for _, r in valid[valid["_rank"].isna()].iterrows():
        if _blank(r["Revision"]):
            add("Warning", "Missing revision", r, "Revision cell is empty.", "No revision")
        else:
            add("Warning", "Revision not understood", r,
                f"Revision '{_txt(r['Revision'])}' could not be ranked - latest may be wrong.", "Check revision")

    for _, r in valid[valid["_floor"] == ""].iterrows():
        add("Warning", "Missing floor", r, "Floor cell is empty.", "No floor")

    for key, g in valid.groupby("_key"):
        for col, label, tag in (("_floor", "Floor", "Floor mismatch"), ("_cat", "Category", "Category mismatch")):
            vals = [v for v in g[col].unique() if v]
            if len(vals) > 1:
                names = sorted({_txt(x) for x in g["Floor" if col == "_floor" else "Category"] if _txt(x)})
                for _, r in g.iterrows():
                    flags.setdefault(int(r["_row"]), [])
                    if tag not in flags[int(r["_row"])]:
                        flags[int(r["_row"])].append(tag)
                add("Warning", f"Same drawing no with different {label.lower()}", g.iloc[0],
                    f"{_txt(g.iloc[0]['Drawing No'])} appears under: {', '.join(names)}", tag)

    exp_norm = [_norm(f) for f in expected_floors]
    for fn, g in d[d["_floor"] != ""].groupby("_floor", sort=False):
        if fn not in exp_norm:
            add("Warning", "Floor not in expected list", g.iloc[0],
                f"'{_txt(g.iloc[0]['Floor'])}' used in {len(g)} row(s) - spelling/naming difference?", "Floor?")

    # --- latest revision per drawing number
    order = valid.sort_values(["_key", "_rank", "_date", "_row"], na_position="first", kind="stable")
    latest_idx = list(order.groupby("_key").tail(1).index)
    d["Status"] = "No drawing no"
    d.loc[valid.index, "Status"] = "Superseded"
    d.loc[latest_idx, "Status"] = "Latest"
    d["Revisions on Record"] = d["_key"].map(valid.groupby("_key").size()).fillna(0).astype(int)

    def date_text(r):
        return r["_date"].strftime("%d-%b-%Y") if pd.notna(r["_date"]) else _txt(r["Date"])

    full = pd.DataFrame({
        "Status": d["Status"], "Drawing No": d["Drawing No"].map(_txt), "Title": d["Title"].map(_txt),
        "Floor": d["Floor"].map(_txt), "Category": d["Category"].map(_txt), "Revision": d["Revision"].map(_txt),
        "Date": d.apply(date_text, axis=1),
        "Flags": d["_row"].map(lambda r: ", ".join(flags.get(int(r), []))),
        "Register Row": d["_row"].astype(int),
    })

    # --- latest register, sorted by expected floor order
    floor_order = {f: i for i, f in enumerate(exp_norm)}
    d["_fo"] = d["_floor"].map(lambda f: floor_order.get(f, 999))
    lat_d = d[d["Status"] == "Latest"].sort_values(["_fo", "_cat", "_key"])
    latest = pd.DataFrame({
        "Drawing No": lat_d["Drawing No"].map(_txt), "Title": lat_d["Title"].map(_txt),
        "Floor": lat_d["Floor"].map(_txt), "Category": lat_d["Category"].map(_txt),
        "Latest Revision": lat_d["Revision"].map(_txt), "Date": full.loc[lat_d.index, "Date"],
        "Revisions on Record": lat_d["Revisions on Record"],
    }).reset_index(drop=True)

    # --- floor x category matrix (latest drawings only)
    counts = d[d["Status"] == "Latest"].groupby(["_floor", "_cat"]).size()
    matrix = pd.DataFrame(0, index=expected_floors, columns=required_cats)
    for f in expected_floors:
        for c in required_cats:
            matrix.loc[f, c] = int(counts.get((_norm(f), _norm(c)), 0))
    missing = [(f, c) for f in expected_floors for c in required_cats if matrix.loc[f, c] == 0]

    iss = pd.DataFrame(issues, columns=["Severity", "Issue", "Drawing No", "Floor", "Register Row", "Details"])
    if len(iss):
        iss = iss.sort_values("Severity", key=lambda s: s.map({"Error": 0, "Warning": 1}), kind="stable")
    iss = iss.reset_index(drop=True)

    stats = {
        "entries": len(full), "unique": int((full["Status"] == "Latest").sum()),
        "superseded": int((full["Status"] == "Superseded").sum()),
        "errors": int((iss["Severity"] == "Error").sum()), "warnings": int((iss["Severity"] == "Warning").sum()),
        "missing": len(missing), "combos": len(expected_floors) * len(required_cats),
    }
    return {"full": full, "latest": latest, "matrix": matrix, "missing": missing, "issues": iss, "stats": stats}


# ------------------------------------------------------------ Excel report
def to_excel(out: dict, source_name: str) -> bytes:
    full, latest, matrix, iss = out["full"], out["latest"], out["matrix"], out["issues"]
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame().to_excel(xw, sheet_name="Summary", index=False)
        latest.to_excel(xw, sheet_name="Latest Register", index=False)
        full.to_excel(xw, sheet_name="Full Register", index=False)
        matrix.reset_index(names="Floor").to_excel(xw, sheet_name="Missing Matrix", index=False)
        iss.to_excel(xw, sheet_name="Issues", index=False)
    buf.seek(0)
    wb = load_workbook(buf)

    def plain(ws):
        _style_header(ws, 1, ws.max_column)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name=FONT, size=10)
                c.border = BORDER
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        _autowidth(ws)

    plain(wb["Latest Register"])

    wf = wb["Full Register"]
    plain(wf)
    for r in range(2, wf.max_row + 1):
        status, flg = wf.cell(r, 1).value, str(wf.cell(r, 8).value or "")
        fill = ORANGE if "Duplicate" in flg else GREY if status == "Superseded" else YELLOW if flg else None
        if status == "Latest" and not flg:
            fill = None
        if fill:
            for c in range(1, wf.max_column + 1):
                wf.cell(r, c).fill = fill

    wi = wb["Issues"]
    plain(wi)
    for r in range(2, wi.max_row + 1):
        fill = RED if wi.cell(r, 1).value == "Error" else YELLOW
        wi.cell(r, 1).fill = fill

    wm = wb["Missing Matrix"]
    plain(wm)
    wm.auto_filter.ref = None
    nr, nc = wm.max_row, wm.max_column
    for r in range(2, nr + 1):
        wm.cell(r, 1).font = Font(name=FONT, size=10, bold=True)
        for c in range(2, nc + 1):
            cell = wm.cell(r, c)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = RED if cell.value == 0 else GREEN
    wm.cell(nr + 2, 1, "Red (0) = no latest drawing found for this floor/category. Green = drawing(s) present.").font = \
        Font(name=FONT, italic=True, color="7F7F7F")

    # --- Summary with live formulas
    sm = wb["Summary"]
    wb.move_sheet("Summary", offset=-4)
    fl, il = max(wf.max_row, 2), max(wi.max_row, 2)
    sm["A1"] = "Drawing Register Check - Summary"
    sm["A1"].font = Font(name=FONT, size=14, bold=True)
    sm["A3"], sm["B3"] = "Source file", source_name
    sm["A3"].font = Font(name=FONT, bold=True)
    sm["B3"].font = Font(name=FONT)
    sm["A5"], sm["B5"] = "Item", "Count"
    _style_header(sm, 5, 2)
    items = [
        ("Total register entries", f"=COUNT('Full Register'!I2:I{fl})"),
        ("Unique drawings (latest revision)", f"=COUNTIF('Full Register'!A2:A{fl},\"Latest\")"),
        ("Superseded (older revisions)", f"=COUNTIF('Full Register'!A2:A{fl},\"Superseded\")"),
        ("Errors (duplicates, missing numbers)", f"=COUNTIF(Issues!A2:A{il},\"Error\")"),
        ("Warnings (revision/floor checks)", f"=COUNTIF(Issues!A2:A{il},\"Warning\")"),
        ("Missing floor/category drawings",
         f"=COUNTIF('Missing Matrix'!B2:{wm.cell(1, nc).column_letter}{nr},0)"),
    ]
    for i, (lbl, f) in enumerate(items):
        sm.cell(6 + i, 1, lbl).font = Font(name=FONT)
        sm.cell(6 + i, 2, f).font = Font(name=FONT)
        sm.cell(6 + i, 1).border = sm.cell(6 + i, 2).border = BORDER
    sm["A13"] = "Tool flags possible issues only - final drawing control by the responsible engineer."
    sm["A13"].font = Font(name=FONT, italic=True, color="7F7F7F")
    sm.column_dimensions["A"].width = 42
    sm.column_dimensions["B"].width = 36

    wb.active = 0
    o = BytesIO()
    wb.save(o)
    return o.getvalue()
