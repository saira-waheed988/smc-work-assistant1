"""Drawing Manager page (Phase 2)."""
import streamlit as st

import area_logic as al
import drawing_logic as dl

NONE = "— none —"
GUESS = {
    "Drawing No": ["drawing no", "dwg no", "drawing number", "dwg", "number", "no"],
    "Title": ["title", "description", "name"],
    "Floor": ["floor", "level"],
    "Category": ["system", "category", "discipline", "type"],
    "Revision": ["rev"],
    "Date": ["date"],
}


def _guess(cols, field):
    for kw in GUESS[field]:
        for i, c in enumerate(cols):
            if kw in c.lower():
                return i + 1
    return 0


def render():
    st.title("Drawing Manager — Drawing Register Check")
    st.write("Apni drawing register (Excel) upload karein. Tool latest revision, duplicates, missing floor drawings "
             "aur galat entries check karke report banayega.")

    f = st.file_uploader("Drawing register (.xlsx)", type=["xlsx", "xlsm"])
    if not f:
        st.info("Register mein kam az kam ye columns hon: Drawing No, Floor, Revision. "
                "Title, System/Category aur Date optional hain.")
        st.stop()

    c1, c2 = st.columns(2)
    sheet = c1.selectbox("Sheet", al.list_sheets(f))
    hrow = c2.number_input("Header row number", 1, 50, 1)
    raw = dl.read_register(f, sheet, hrow)
    cols = [c for c in raw.columns if c != "_row"]
    if not cols:
        st.error("Sheet khali hai ya header row galat hai.")
        st.stop()

    with st.expander("🔗 Column mapping (kaun sa column kya hai)", expanded=True):
        m = st.columns(3)
        mapping = {}
        for i, field in enumerate(dl.STD):
            sel = m[i % 3].selectbox(field, [NONE] + cols, index=_guess(cols, field), key=f"map_{f.name}_{field}")
            mapping[field] = None if sel == NONE else sel
    for need in ("Drawing No", "Floor", "Revision"):
        if not mapping[need]:
            st.error(f"'{need}' column select karna zaroori hai.")
            st.stop()

    reg = dl.standardize(raw, mapping)
    found_floors = dl.found_values(reg, "Floor")
    found_cats = dl.found_values(reg, "Category")

    with st.expander("🏢 Expected floors aur required systems", expanded=True):
        txt = st.text_area("Expected floors (har line mein ek floor, building ke order mein)",
                           "\n".join(found_floors), height=140, key=f"floors_{f.name}_{mapping['Floor']}")
        expected = [x.strip() for x in txt.splitlines() if x.strip()]
        required = st.multiselect("Har floor par kaun si drawing categories zaroori hain?",
                                  found_cats, default=found_cats, key=f"cats_{f.name}_{mapping['Category']}")
    if not expected or not required:
        st.warning("Kam az kam ek floor aur ek category chahiye.")
        st.stop()

    out = dl.analyze(reg, expected, required)
    s = out["stats"]
    k = st.columns(5)
    k[0].metric("Register entries", s["entries"])
    k[1].metric("Unique drawings", s["unique"])
    k[2].metric("Superseded", s["superseded"])
    k[3].metric("🔴 Errors", s["errors"])
    k[4].metric("Missing drawings", f"{s['missing']} / {s['combos']}")

    tab1, tab2, tab3, tab4 = st.tabs(["Latest register", "Full register", "Missing matrix", "Issues"])
    with tab1:
        st.dataframe(out["latest"], use_container_width=True, hide_index=True)
    with tab2:
        def colour(r):
            fl = r["Flags"]
            bg = "#F4B183" if "Duplicate" in fl else "#E7E6E6" if r["Status"] == "Superseded" \
                else "#FFEB9C" if fl else ""
            return [f"background-color: {bg}"] * len(r)
        st.dataframe(out["full"].style.apply(colour, axis=1), use_container_width=True, hide_index=True)
        st.caption("Orange = duplicate · Grey = purani revision · Yellow = warning")
    with tab3:
        mx = out["matrix"]
        styler = mx.style
        paint = lambda v: "background-color: #FFC7CE" if v == 0 else "background-color: #C6EFCE"
        styler = styler.map(paint) if hasattr(styler, "map") else styler.applymap(paint)
        st.dataframe(styler, use_container_width=True)
        if out["missing"]:
            st.error("Missing: " + "; ".join(f"{fl} → {c}" for fl, c in out["missing"]))
        else:
            st.success("Har floor par sari required categories ki drawings maujood hain.")
    with tab4:
        if len(out["issues"]):
            st.dataframe(out["issues"], use_container_width=True, hide_index=True)
        else:
            st.success("Koi issue nahi mila.")

    st.download_button("⬇️ Download Excel report", dl.to_excel(out, f.name), "Drawing_Register_Report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
