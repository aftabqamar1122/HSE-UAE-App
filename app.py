import streamlit as st
from dotenv import load_dotenv
import os

load_dotenv()

# ── Page config ──────────────────────────────────────────
st.set_page_config(
    page_title="UAE HSE Compliance App",
    page_icon="🦺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Emirate Selector (MANDATORY FIRST SCREEN) ─────────────
st.title("🦺 UAE Construction HSE Compliance App")
st.markdown("**Abu Dhabi & Dubai | ADOSH-SF v4.0 | Dubai Municipality Code**")
st.divider()

if "emirate" not in st.session_state:
    st.session_state.emirate = None

col1, col2 = st.columns(2)

with col1:
    if st.button("🔵 ABU DHABI", use_container_width=True, type="primary"):
        st.session_state.emirate = "Abu Dhabi"

with col2:
    if st.button("🔴 DUBAI", use_container_width=True, type="secondary"):
        st.session_state.emirate = "Dubai"

if st.session_state.emirate:
    st.success(f"✅ Selected Emirate: **{st.session_state.emirate}**")
    st.info("👈 Use the sidebar to navigate between modules")
    
    st.markdown("### Available Modules")
    st.markdown("""
    | Module | Description |
    |---|---|
    | 📚 Knowledge Base | Browse all CoPs (Abu Dhabi) or Code Chapters (Dubai) |
    | 🤖 AI Chatbot | Ask any HSE question — gets Emirates-specific answers |
    | 📋 Risk Assessment | Generate HIRA/JSA templates |
    | 🗣️ Toolbox Talks | Generate and print toolbox talks |
    | 📝 Permit to Work | Fill and save digital permits |
    | 🚨 Incident Report | Log and submit incident reports |
    """)
else:
    st.warning("⬆️ Please select your Emirate above to begin")