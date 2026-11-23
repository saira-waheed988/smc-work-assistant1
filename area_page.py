"""Area Tracker page (Phase 1)."""
import pandas as pd
import streamlit as st

import area_logic as al


def render():
    st.title("Area Tracker — Excel Comparison")
    st.write("Purani aur nayi area list upload karein. Tool New / Removed / Modified / Unchanged areas dhoondh kar "
             "colour-coded Excel report banayega.")

    c1, c2 = st.columns(2)
    with c1:
        f_old = st.file_uploader("Previous area list (.xlsx)", type=["xlsx", "xlsm"], key="old")
    with c2:
        f_new = st.file_uploader("Updated area list (.xlsx)", type=["xlsx", "xlsm"], key="new")

    if not (f_old and f_new):
        st.stop()

    with st.expander("⚙️ Sheet / header settings", expanded=True):
        s1, s2 = st.columns(2)
        sheet_old = s1.selectbox("Sheet (Previous)", al.list_sheets(f_old))
        sheet_new = s2.selectbox("Sheet (Updated)", al.list_sheets(f_new))
        hrow = st.number_input("Header row number (column names kis row mein hain)", 1, 50, 1)

    old_df = al.read_sheet(f_old, sheet_old, hrow)
    new_df = al.read_sheet(f_new, sheet_new, hrow)
    common = [c for c in new_df.columns if c in old_df.columns]
    if not common:
        st.error("Dono sheets mein koi common column nahi mila. Header row check karein.")
        st.stop()

    with st.expander("🔑 Comparison settings", expanded=True):
        key = st.selectbox("Key column (unique ID jaise Room No / Space ID)", common)
        num_cols = [c for c in common if pd.to_numeric(new_df[c], errors="coerce").notna().mean() > 0.6]
        area_col = st.selectbox("Area column (totals ke liye, optional)", ["— none —"] + num_cols)
        area_col = None if area_col == "— none —" else area_col
        ignore = st.multiselect("Columns jinhe compare NAHI karna", [c for c in common if c != key])
        tol = st.number_input("Numeric tolerance (e.g. 0.01 m²)", 0.0, 100.0, 0.0, 0.01, format="%.3f")

    out = al.compare(old_df, new_df, key, tol, ignore)
    res = out["result"]
    cnt = al.summary_counts(res)

    m = st.columns(4)
    m[0].metric("🟢 New", cnt["New"])
    m[1].metric("🟡 Modified", cnt["Modified"])
    m[2].metric("🔴 Removed", cnt["Removed"])
    m[3].metric("⚪ Unchanged", cnt["Unchanged"])

    t = al.totals(res, area_col)
    if t:
        a = st.columns(3)
        a[0].metric("Added area", f"{t['added']:,.2f}")
        a[1].metric("Removed area", f"{t['removed']:,.2f}")
        a[2].metric("Updated list total", f"{t['updated_total']:,.2f}")

    for n in out["notes"]:
        st.warning(n)

    tab1, tab2 = st.tabs(["Comparison", "Changes detail"])
    COLORS = {"New": "#C6EFCE", "Removed": "#FFC7CE", "Modified": "#FFEB9C"}
    with tab1:
        show = st.multiselect("Status filter", list(cnt), default=["New", "Modified", "Removed"])
        view = res[res["Status"].isin(show)]
        st.dataframe(view.style.apply(
            lambda r: [f"background-color: {COLORS.get(r['Status'], '')}"] * len(r), axis=1),
            use_container_width=True, hide_index=True)
    with tab2:
        if len(out["changes"]):
            st.dataframe(out["changes"], use_container_width=True, hide_index=True)
        else:
            st.write("Koi value change nahi hui.")

    report = al.to_excel(out, key, area_col, f_old.name, f_new.name)
    st.download_button("⬇️ Download Excel report", report, "Area_Comparison_Report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
