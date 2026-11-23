"""SMC Work Assistant.  Run:  streamlit run app.py"""
import streamlit as st

st.set_page_config(page_title="SMC Work Assistant", page_icon="📐", layout="wide")

with st.sidebar:
    st.title("📐 SMC Work Assistant")
    page = st.radio("Module", [
        "01 · Area Tracker",
        "02 · Drawing Manager",
        "03 · ELV Calculator (Phase 3)",
        "04 · QA Checker (Phase 5)",
    ])
    st.caption("Confidential project files ko company permission ke baghair upload na karein.")

if page.startswith("01"):
    import area_page
    area_page.render()
elif page.startswith("02"):
    import drawing_page
    drawing_page.render()
else:
    st.title(page.split("·")[1].strip())
    st.info("Ye module aglay phase mein banega.")
