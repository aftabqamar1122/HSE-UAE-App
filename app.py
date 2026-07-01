import streamlit as st

# ── MUST BE FIRST — before any other st. calls ────────────
st.set_page_config(
    page_title="UAE HSE Compliance App — ENGC",
    page_icon="🦺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── NOW IMPORT AUTH ───────────────────────────────────────
from auth import (
    check_access,
    show_login_page,
    show_paywall,
    show_trial_banner,
    trial_expired,
)
import time
from dotenv import load_dotenv
import os
from datetime import date

load_dotenv()

# ── ACCESS CONTROL ────────────────────────────────────────
if st.session_state.get("show_paywall"):
    show_paywall()
    st.stop()

if (st.session_state.get("trial_started")
        and trial_expired()):
    show_paywall()
    st.stop()

if not check_access():
    show_login_page()
    st.stop()

if (st.session_state.get("trial_started")
        and not st.session_state.get(
            "authenticated")):
    show_trial_banner()

# ── YOUR EXISTING app.py CODE CONTINUES BELOW ─────────────
# (Keep all your existing CSS and content below this line)
# (Do NOT add another st.set_page_config # ── CUSTOM SIDEBAR CSS ────────────────────────────────────
st.markdown("""
<style>

/* ── Sidebar background ── */
[data-testid="stSidebar"] {
    background: linear-gradient(
        180deg,
        #1F3864 0%,
        #16294d 40%,
        #0d1a33 100%
    ) !important;
}

/* ── Sidebar width ── */
[data-testid="stSidebar"] {
    min-width: 260px !important;
    max-width: 260px !important;
}

/* ── All sidebar text white ── */
[data-testid="stSidebar"] * {
    color: white !important;
}

/* ── Sidebar nav links ── */
[data-testid="stSidebarNav"] a {
    display: block !important;
    padding: 10px 16px !important;
    margin: 3px 8px !important;
    border-radius: 8px !important;
    color: rgba(255,255,255,0.85) !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    text-decoration: none !important;
    transition: all 0.2s ease !important;
    border-left: 3px solid transparent !important;
}

/* ── Sidebar nav hover ── */
[data-testid="stSidebarNav"] a:hover {
    background: rgba(255,255,255,0.12) !important;
    border-left: 3px solid #FFC000 !important;
    color: white !important;
    padding-left: 20px !important;
}

/* ── Active/selected nav link ── */
[data-testid="stSidebarNav"] a[aria-current="page"] {
    background: rgba(255,255,255,0.18) !important;
    border-left: 3px solid #C00000 !important;
    color: white !important;
    font-weight: 700 !important;
}

/* ── Sidebar buttons ── */
[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,255,255,0.12) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
    border-radius: 8px !important;
    width: 100% !important;
    font-size: 13px !important;
    padding: 8px !important;
    transition: all 0.2s !important;
}

[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.22) !important;
    border-color: #FFC000 !important;
    color: #FFC000 !important;
}

/* ── Sidebar selectbox ── */
[data-testid="stSidebar"] .stSelectbox > div > div {
    background: rgba(255,255,255,0.1) !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
    color: white !important;
    border-radius: 8px !important;
}

/* ── Sidebar divider ── */
[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,0.15) !important;
    margin: 8px 0 !important;
}

/* ── Main page background ── */
.main {
    background-color: #f4f6f9 !important;
}

/* ── Hide default Streamlit menu ── */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }

</style>
""", unsafe_allow_html=True)
# ── SIDEBAR ───────────────────────────────────────────────
with st.sidebar:

    # ── App Logo / Title ──────────────────────────────
    st.markdown("""
<div style="
    text-align: center;
    padding: 16px 8px 8px 8px;
">
    <div style="font-size: 40px;">🦺</div>
    <div style="
        font-size: 15px;
        font-weight: 800;
        color: white;
        letter-spacing: 0.5px;
        margin-top: 6px;
        line-height: 1.3;
    ">
        UAE HSE<br>Compliance App
    </div>
    <div style="
        font-size: 10px;
        color: rgba(255,255,255,0.6);
        margin-top: 4px;
    ">
        ADOSH-SF v4.0 | DM Code
    </div>
</div>
""", unsafe_allow_html=True)

    st.markdown(
        "<hr>", unsafe_allow_html=True
    )

    # ── Emirate Selector in Sidebar ───────────────────
    st.markdown("""
<div style="
    font-size: 11px;
    font-weight: 700;
    color: rgba(255,255,255,0.6);
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 0 4px;
    margin-bottom: 6px;
">
    🗺️ Select Emirate
</div>
""", unsafe_allow_html=True)

    if "emirate" not in st.session_state:
        st.session_state.emirate = None

    col_sb1, col_sb2 = st.columns(2)
    with col_sb1:
        if st.button(
            "🔵 Abu\nDhabi",
            key="sb_ad",
            use_container_width=True
        ):
            st.session_state.emirate = "Abu Dhabi"
            st.rerun()

    with col_sb2:
        if st.button(
            "🔴\nDubai",
            key="sb_dxb",
            use_container_width=True
        ):
            st.session_state.emirate = "Dubai"
            st.rerun()

    # Show selected emirate
    if st.session_state.emirate == "Abu Dhabi":
        st.markdown("""
<div style="
    background: rgba(0,100,200,0.25);
    border: 1px solid rgba(100,180,255,0.4);
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 12px;
    color: #90CAF9;
    text-align: center;
    margin: 6px 0;
">
    ✅ <strong>Abu Dhabi</strong> Active<br>
    <span style="font-size:10px;
                 color:rgba(255,255,255,0.5);">
        ADOSH-SF v4.0 | ADPHC
    </span>
</div>
""", unsafe_allow_html=True)

    elif st.session_state.emirate == "Dubai":
        st.markdown("""
<div style="
    background: rgba(200,0,0,0.25);
    border: 1px solid rgba(255,100,100,0.4);
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 12px;
    color: #EF9A9A;
    text-align: center;
    margin: 6px 0;
">
    ✅ <strong>Dubai</strong> Active<br>
    <span style="font-size:10px;
                 color:rgba(255,255,255,0.5);">
        Dubai Municipality Code
    </span>
</div>
""", unsafe_allow_html=True)

    else:
        st.markdown("""
<div style="
    background: rgba(255,200,0,0.15);
    border: 1px solid rgba(255,200,0,0.3);
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 11px;
    color: #FFC000;
    text-align: center;
    margin: 6px 0;
">
    ⬆️ Select an Emirate above
</div>
""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Navigation Label ──────────────────────────────
    st.markdown("""
<div style="
    font-size: 11px;
    font-weight: 700;
    color: rgba(255,255,255,0.6);
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 0 4px;
    margin-bottom: 6px;
">
    📱 Modules
</div>
""", unsafe_allow_html=True)

    # Module buttons with icons
    modules = [
        ("📚", "Knowledge Base",    "Knowledge_Base"),
        ("🤖", "AI Chatbot",        "AI_Chatbot"),
        ("📋", "Risk Assessment",   "Risk_Assessment"),
        ("🗣️", "Toolbox Talks",     "Toolbox_Talks"),
        ("📝", "Permit to Work",    "Permit_to_Work"),
        ("🚨", "Incident Report",   "Incident_Report"),
    ]

    for icon, name, page in modules:
        st.markdown(f"""
<div style="
    display: flex;
    align-items: center;
    padding: 9px 12px;
    margin: 3px 0;
    border-radius: 8px;
    background: rgba(255,255,255,0.06);
    border-left: 3px solid rgba(255,255,255,0.15);
    cursor: pointer;
    transition: all 0.2s;
    font-size: 13px;
    color: rgba(255,255,255,0.88);
">
    <span style="font-size:16px;
                 margin-right:10px;">
        {icon}
    </span>
    {name}
</div>
""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Emergency Numbers ─────────────────────────────
    st.markdown("""
<div style="
    font-size: 11px;
    font-weight: 700;
    color: rgba(255,255,255,0.6);
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 0 4px;
    margin-bottom: 8px;
">
    🚨 Emergency Numbers
</div>
<div style="
    background: rgba(192,0,0,0.2);
    border: 1px solid rgba(255,100,100,0.3);
    border-radius: 8px;
    padding: 10px 12px;
    font-size: 12px;
    line-height: 2;
    color: rgba(255,255,255,0.9);
">
    🚔 Police: <strong>999</strong><br>
    🚑 Ambulance: <strong>998</strong><br>
    🚒 Civil Defence: <strong>997</strong><br>
    🏛️ DM Emergency: <strong>800900</strong>
</div>
""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Summer Ban Alert ──────────────────────────────
    today = date.today()
    is_ban = (
        (today.month == 6 and today.day >= 15)
        or today.month in [7, 8]
        or (today.month == 9 and today.day <= 15)
    )

    if is_ban:
        st.markdown("""
<div style="
    background: rgba(255,100,0,0.25);
    border: 1px solid rgba(255,150,0,0.5);
    border-radius: 8px;
    padding: 10px 12px;
    font-size: 11px;
    color: #FFB74D;
    text-align: center;
    line-height: 1.7;
">
    ☀️ <strong>MIDDAY BAN ACTIVE</strong><br>
    No outdoor work<br>
    <strong>12:30 — 15:00</strong><br>
    15 Jun – 15 Sep<br>
    <span style="font-size:10px;
                 color:rgba(255,255,255,0.5);">
        MOHRE Res. 44/2022
    </span>
</div>
""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Developer Info ────────────────────────────────
    st.markdown("""
<div style="
    text-align: center;
    padding: 4px;
    font-size: 11px;
    color: rgba(255,255,255,0.45);
    line-height: 1.7;
">
    Developed by<br>
    <strong style="color:rgba(255,255,255,0.7);">
        Muhammad Aftab Qamar
    </strong><br>
    Senior HSSE Engineer | ENGC<br>
    📱 +971 52 267 1122
</div>
""", unsafe_allow_html=True)
    # ── SIDEBAR ───────────────────────────────────────────────
with st.sidebar:

    # ── App Logo / Title ──────────────────────────────
    st.markdown("""
<div style="
    text-align: center;
    padding: 16px 8px 8px 8px;
">
    <div style="font-size: 40px;">🦺</div>
    <div style="
        font-size: 15px;
        font-weight: 800;
        color: white;
        letter-spacing: 0.5px;
        margin-top: 6px;
        line-height: 1.3;
    ">
        UAE HSE<br>Compliance App
    </div>
    <div style="
        font-size: 10px;
        color: rgba(255,255,255,0.6);
        margin-top: 4px;
    ">
        ADOSH-SF v4.0 | DM Code
    </div>
</div>
""", unsafe_allow_html=True)

    st.markdown(
        "<hr>", unsafe_allow_html=True
    )

    # ── Emirate Selector in Sidebar ───────────────────
    st.markdown("""
<div style="
    font-size: 11px;
    font-weight: 700;
    color: rgba(255,255,255,0.6);
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 0 4px;
    margin-bottom: 6px;
">
    🗺️ Select Emirate
</div>
""", unsafe_allow_html=True)

    if "emirate" not in st.session_state:
        st.session_state.emirate = None

    col_sb1, col_sb2 = st.columns(2)
    with col_sb1:
        if st.button(
            "🔵 Abu\nDhabi",
            key="sb_ad",
            use_container_width=True
        ):
            st.session_state.emirate = "Abu Dhabi"
            st.rerun()

    with col_sb2:
        if st.button(
            "🔴\nDubai",
            key="sb_dxb",
            use_container_width=True
        ):
            st.session_state.emirate = "Dubai"
            st.rerun()

    # Show selected emirate
    if st.session_state.emirate == "Abu Dhabi":
        st.markdown("""
<div style="
    background: rgba(0,100,200,0.25);
    border: 1px solid rgba(100,180,255,0.4);
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 12px;
    color: #90CAF9;
    text-align: center;
    margin: 6px 0;
">
    ✅ <strong>Abu Dhabi</strong> Active<br>
    <span style="font-size:10px;
                 color:rgba(255,255,255,0.5);">
        ADOSH-SF v4.0 | ADPHC
    </span>
</div>
""", unsafe_allow_html=True)

    elif st.session_state.emirate == "Dubai":
        st.markdown("""
<div style="
    background: rgba(200,0,0,0.25);
    border: 1px solid rgba(255,100,100,0.4);
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 12px;
    color: #EF9A9A;
    text-align: center;
    margin: 6px 0;
">
    ✅ <strong>Dubai</strong> Active<br>
    <span style="font-size:10px;
                 color:rgba(255,255,255,0.5);">
        Dubai Municipality Code
    </span>
</div>
""", unsafe_allow_html=True)

    else:
        st.markdown("""
<div style="
    background: rgba(255,200,0,0.15);
    border: 1px solid rgba(255,200,0,0.3);
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 11px;
    color: #FFC000;
    text-align: center;
    margin: 6px 0;
">
    ⬆️ Select an Emirate above
</div>
""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Navigation Label ──────────────────────────────
    st.markdown("""
<div style="
    font-size: 11px;
    font-weight: 700;
    color: rgba(255,255,255,0.6);
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 0 4px;
    margin-bottom: 6px;
">
    📱 Modules
</div>
""", unsafe_allow_html=True)

    # Module buttons with icons
    modules = [
        ("📚", "Knowledge Base",    "Knowledge_Base"),
        ("🤖", "AI Chatbot",        "AI_Chatbot"),
        ("📋", "Risk Assessment",   "Risk_Assessment"),
        ("🗣️", "Toolbox Talks",     "Toolbox_Talks"),
        ("📝", "Permit to Work",    "Permit_to_Work"),
        ("🚨", "Incident Report",   "Incident_Report"),
    ]

    for icon, name, page in modules:
        st.markdown(f"""
<div style="
    display: flex;
    align-items: center;
    padding: 9px 12px;
    margin: 3px 0;
    border-radius: 8px;
    background: rgba(255,255,255,0.06);
    border-left: 3px solid rgba(255,255,255,0.15);
    cursor: pointer;
    transition: all 0.2s;
    font-size: 13px;
    color: rgba(255,255,255,0.88);
">
    <span style="font-size:16px;
                 margin-right:10px;">
        {icon}
    </span>
    {name}
</div>
""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Emergency Numbers ─────────────────────────────
    st.markdown("""
<div style="
    font-size: 11px;
    font-weight: 700;
    color: rgba(255,255,255,0.6);
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 0 4px;
    margin-bottom: 8px;
">
    🚨 Emergency Numbers
</div>
<div style="
    background: rgba(192,0,0,0.2);
    border: 1px solid rgba(255,100,100,0.3);
    border-radius: 8px;
    padding: 10px 12px;
    font-size: 12px;
    line-height: 2;
    color: rgba(255,255,255,0.9);
">
    🚔 Police: <strong>999</strong><br>
    🚑 Ambulance: <strong>998</strong><br>
    🚒 Civil Defence: <strong>997</strong><br>
    🏛️ DM Emergency: <strong>800900</strong>
</div>
""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Summer Ban Alert ──────────────────────────────
    today = date.today()
    is_ban = (
        (today.month == 6 and today.day >= 15)
        or today.month in [7, 8]
        or (today.month == 9 and today.day <= 15)
    )

    if is_ban:
        st.markdown("""
<div style="
    background: rgba(255,100,0,0.25);
    border: 1px solid rgba(255,150,0,0.5);
    border-radius: 8px;
    padding: 10px 12px;
    font-size: 11px;
    color: #FFB74D;
    text-align: center;
    line-height: 1.7;
">
    ☀️ <strong>MIDDAY BAN ACTIVE</strong><br>
    No outdoor work<br>
    <strong>12:30 — 15:00</strong><br>
    15 Jun – 15 Sep<br>
    <span style="font-size:10px;
                 color:rgba(255,255,255,0.5);">
        MOHRE Res. 44/2022
    </span>
</div>
""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Developer Info ────────────────────────────────
    st.markdown("""
<div style="
    text-align: center;
    padding: 4px;
    font-size: 11px;
    color: rgba(255,255,255,0.45);
    line-height: 1.7;
">
    Developed by<br>
    <strong style="color:rgba(255,255,255,0.7);">
        Muhammad Aftab Qamar
    </strong><br>
    Senior HSSE Engineer | ENGC<br>
    📱 +971 52 267 1122
</div>
""", unsafe_allow_html=True)
import streamlit as st
from dotenv import load_dotenv
import os
from datetime import date

load_dotenv()

st.set_page_config(
    page_title="UAE HSE Compliance App — ENGC",
    page_icon="🦺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CUSTOM CSS ────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background */
    .main { background-color: #f8f9fa; }

    /* Project card style */
    .project-card {
        background: white;
        border-left: 5px solid #1F3864;
        border-radius: 8px;
        padding: 12px 16px;
        margin: 6px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }
    .project-card:hover {
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        border-left-color: #C00000;
    }
    .project-name {
        font-weight: 700;
        font-size: 14px;
        color: #1F3864;
    }
    .project-type {
        font-size: 11px;
        color: #666;
        margin-top: 2px;
    }

    /* Contact card */
    .contact-card {
        background: linear-gradient(
            135deg, #1F3864 0%, #2E4F8A 100%
        );
        border-radius: 12px;
        padding: 20px;
        color: white;
        text-align: center;
    }
    .contact-name {
        font-size: 20px;
        font-weight: 700;
        margin-bottom: 4px;
    }
    .contact-title {
        font-size: 13px;
        opacity: 0.85;
        margin-bottom: 12px;
    }
    .contact-detail {
        font-size: 13px;
        margin: 4px 0;
        opacity: 0.95;
    }

    /* Header banner */
    .header-banner {
        background: linear-gradient(
            135deg, #1F3864 0%, #C00000 100%
        );
        border-radius: 12px;
        padding: 28px 32px;
        color: white;
        margin-bottom: 20px;
    }
    .app-title {
        font-size: 32px;
        font-weight: 800;
        margin: 0;
        letter-spacing: 1px;
    }
    .app-subtitle {
        font-size: 15px;
        opacity: 0.9;
        margin-top: 6px;
    }

    /* Emirate buttons */
    .stButton > button {
        height: 70px;
        font-size: 18px;
        font-weight: 700;
        border-radius: 10px;
        border: none;
        width: 100%;
        transition: all 0.2s;
    }

    /* Info boxes */
    .info-box {
        background: #EEF2FF;
        border-radius: 8px;
        padding: 10px 14px;
        font-size: 13px;
        color: #1F3864;
        margin: 4px 0;
    }

    /* Section headers */
    .section-hdr {
        background: #1F3864;
        color: white;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 14px;
        margin: 12px 0 8px 0;
    }

    /* Module pills */
    .module-pill {
        display: inline-block;
        background: #E8F0FE;
        border: 1px solid #1F3864;
        border-radius: 20px;
        padding: 4px 12px;
        font-size: 12px;
        color: #1F3864;
        font-weight: 600;
        margin: 3px;
    }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# HEADER BANNER
# ═══════════════════════════════════════════════════════════
st.markdown("""
<div class="header-banner">
    <div class="app-title">
        🦺 UAE Construction HSE Compliance App
    </div>
    <div class="app-subtitle">
        Abu Dhabi & Dubai | ADOSH-SF v4.0 |
        Dubai Municipality Code |
        AI-Powered HSE Management System
    </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# ROW 1: CONTACT + EMIRATE SELECTOR
# ═══════════════════════════════════════════════════════════
col_contact, col_selector = st.columns([1, 1.4])

# ── CONTACT CARD ──────────────────────────────────────────
with col_contact:
    st.markdown("""
<div class="contact-card">
    <div class="contact-name">
        Muhammad Aftab Qamar
    </div>
    <div class="contact-title">
        Senior HSSE Engineer<br>
        MUHAMMAD AFTAB QAMAR (IDIP NEBOSH)
    </div>
    <hr style="border-color:rgba(255,255,255,0.3);
               margin:10px 0;">
    <div class="contact-detail">
        📱 +971 522671122
    </div>
    <div class="contact-detail">
        ✉️ aftabqamar1122@gmail.com
    </div>
    <div class="contact-detail">
        📍 Khalifa City, Abu Dhabi, UAE
    </div>
    <hr style="border-color:rgba(255,255,255,0.3);
               margin:10px 0;">
    <div style="font-size:11px; opacity:0.8;">
       15 YEARS + EXPERIENCE| MASTER IN QPM & MEd | ADOSH SENIOR PRACTISIONER | NEBOSH Idip | ISO 45001 | ADOSH-SF v4.0
    </div>
</div>
""", unsafe_allow_html=True)

# ── EMIRATE SELECTOR ──────────────────────────────────────
with col_selector:
    st.markdown(
        '<div class="section-hdr">'
        '🗺️ SELECT YOUR EMIRATE TO BEGIN'
        '</div>',
        unsafe_allow_html=True
    )

    if "emirate" not in st.session_state:
        st.session_state.emirate = None

    c1, c2 = st.columns(2)
    with c1:
        if st.button(
            "🔵\nABU DHABI\nADOSH-SF v4.0",
            use_container_width=True,
            type="primary",
            key="btn_ad"
        ):
            st.session_state.emirate = "Abu Dhabi"

    with c2:
        if st.button(
            "🔴\nDUBAI\nDM Code",
            use_container_width=True,
            type="secondary",
            key="btn_dxb"
        ):
            st.session_state.emirate = "Dubai"

    if st.session_state.emirate:
        if st.session_state.emirate == "Abu Dhabi":
            st.success(
                f"✅ **Abu Dhabi** selected — "
                f"ADOSH-SF v4.0 framework active"
            )
        else:
            st.error(
                f"✅ **Dubai** selected — "
                f"Dubai Municipality Code active"
            )
        st.info(
            "👈 Use the sidebar to navigate modules"
        )
    else:
        st.warning(
            "⬆️ Please select your Emirate above "
            "to activate the app"
        )

    st.divider()

    # Quick emergency reference
    st.markdown(
        '<div class="section-hdr">'
        '🚨 UAE EMERGENCY NUMBERS'
        '</div>',
        unsafe_allow_html=True
    )
    em1, em2, em3, em4 = st.columns(4)
    em1.metric("Police", "999")
    em2.metric("Ambulance", "998")
    em3.metric("Civil Defence", "997")
    em4.metric("DM Emergency", "800900")

st.divider()

# ═══════════════════════════════════════════════════════════
# ROW 2: PROJECTS + APP MODULES
# ═══════════════════════════════════════════════════════════
col_projects, col_modules = st.columns([1.3, 1])

# ── PROJECTS PANEL ────────────────────────────────────────
with col_projects:
    st.markdown(
        '<div class="section-hdr">'
        '🏗️ BLOOM LIVING PROJECTS — ABU DHABI'
        '</div>',
        unsafe_allow_html=True
    )

    # Project Info Row
    pi1, pi2, pi3 = st.columns(3)
    with pi1:
        st.markdown("""
<div class="info-box">
    <strong>CLIENT</strong><br>
    Bloom Living / Bloom District Properties
</div>""", unsafe_allow_html=True)
    with pi2:
        st.markdown("""
<div class="info-box">
    <strong>PMC / TIMES</strong><br>
    Times Development (PMC)
</div>""", unsafe_allow_html=True)
    with pi3:
        st.markdown("""
<div class="info-box">
    <strong>CONSULTANT</strong><br>
    Dar Al-Handasah
</div>""", unsafe_allow_html=True)

    pi4, pi5 = st.columns(2)
    with pi4:
        st.markdown("""
<div class="info-box">
    <strong>MAIN CONTRACTOR</strong><br>
    Exeed National General Contracting LLC (ENGC)
</div>""", unsafe_allow_html=True)
    with pi5:
        st.markdown("""
<div class="info-box">
    <strong>SUB-CONTRACTOR</strong><br>
    ___________________________
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── VILLA PROJECTS ─────────────────────────────────
    st.markdown(
        '<div class="section-hdr" '
        'style="background:#8B0000;">'
        '🏡 RESIDENTIAL VILLA PROJECTS'
        '</div>',
        unsafe_allow_html=True
    )

    villa_projects = [
        {
            "name": "Olvera — Bloom Living",
            "type": "Residential Villa Development",
            "ref": "BL-VL-01",
        },
        {
            "name": "Almeria — Bloom Living",
            "type": "Residential Villa Development",
            "ref": "BL-VL-02",
        },
    ]

    for p in villa_projects:
        st.markdown(f"""
<div class="project-card">
    <div class="project-name">
        🏡 {p['name']}
    </div>
    <div class="project-type">
        {p['type']} &nbsp;|&nbsp;
        Ref: {p['ref']} &nbsp;|&nbsp;
        Abu Dhabi, UAE
    </div>
</div>""", unsafe_allow_html=True)

    # ── COMMUNITY CENTER PROJECTS ───────────────────────
    st.markdown(
        '<div class="section-hdr" '
        'style="background:#1F3864; margin-top:12px;">'
        '🏛️ COMMUNITY CENTER PROJECTS'
        '</div>',
        unsafe_allow_html=True
    )

    cc_projects = [
        {
            "name": "Cordoba Community Center",
            "type": "Community Center — G+1",
            "ref": "BL-CC-01",
        },
        {
            "name": "Toledo Community Center",
            "type": "Community Center — G+1",
            "ref": "BL-CC-02",
        },
        {
            "name": "Casares Community Center",
            "type": "Community Center — G+1",
            "ref": "BL-CC-03",
        },
    ]

    for p in cc_projects:
        st.markdown(f"""
<div class="project-card"
     style="border-left-color:#1F6864;">
    <div class="project-name"
         style="color:#1F6864;">
        🏛️ {p['name']}
    </div>
    <div class="project-type">
        {p['type']} &nbsp;|&nbsp;
        Ref: {p['ref']} &nbsp;|&nbsp;
        Abu Dhabi, UAE
    </div>
</div>""", unsafe_allow_html=True)

    # ── SERVICE STATION ─────────────────────────────────
    st.markdown(
        '<div class="section-hdr" '
        'style="background:#5C4033; margin-top:12px;">'
        '⛽ SERVICE STATION PROJECT'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown("""
<div class="project-card"
     style="border-left-color:#5C4033;">
    <div class="project-name"
         style="color:#5C4033;">
        ⛽ Service Station — Bloom Living
    </div>
    <div class="project-type">
        Service Center / Plot C13, Zayed City &nbsp;|&nbsp;
        Ref: BL-SS-01 &nbsp;|&nbsp;
        Abu Dhabi, UAE
    </div>
</div>""", unsafe_allow_html=True)

# ── APP MODULES PANEL ─────────────────────────────────────
with col_modules:
    st.markdown(
        '<div class="section-hdr">'
        '📱 APP MODULES'
        '</div>',
        unsafe_allow_html=True
    )

    modules = [
        {
            "icon": "📚",
            "name": "Knowledge Base",
            "desc": (
                "Browse all 54 ADOSH CoPs or "
                "23 Dubai Code Chapters. "
                "AI-powered search."
            ),
            "color": "#1F3864",
        },
        {
            "icon": "🤖",
            "name": "AI HSE Chatbot",
            "desc": (
                "Ask any HSE question. "
                "Gets Emirates-specific answers "
                "with CoP references."
            ),
            "color": "#0D47A1",
        },
        {
            "icon": "📋",
            "name": "Risk Assessment",
            "desc": (
                "Generate ENGC 14-column HIRA "
                "with hierarchy of controls. "
                "Downloads as .docx."
            ),
            "color": "#1B5E20",
        },
        {
            "icon": "🗣️",
            "name": "Toolbox Talks",
            "desc": (
                "Generate AI-powered toolbox talks "
                "for any HSE topic with "
                "attendance sheet."
            ),
            "color": "#4A148C",
        },
        {
            "icon": "📝",
            "name": "Permit to Work",
            "desc": (
                "Digital PTW forms — Hot Work, "
                "Confined Space, Excavation, "
                "Heights, LOTO, Lifting."
            ),
            "color": "#BF360C",
        },
        {
            "icon": "🚨",
            "name": "Incident Report",
            "desc": (
                "Full incident investigation with "
                "5-Why analysis and corrective "
                "action tables."
            ),
            "color": "#C00000",
        },
    ]

    for mod in modules:
        st.markdown(f"""
<div class="project-card"
     style="border-left-color:{mod['color']};
            margin:5px 0; padding:10px 14px;">
    <div class="project-name"
         style="color:{mod['color']}; font-size:15px;">
        {mod['icon']} &nbsp;{mod['name']}
    </div>
    <div class="project-type" style="margin-top:3px;">
        {mod['desc']}
    </div>
</div>""", unsafe_allow_html=True)

    st.divider()

    # ── REGULATORY FRAMEWORKS ──────────────────────────
    st.markdown(
        '<div class="section-hdr">'
        '⚖️ REGULATORY FRAMEWORKS'
        '</div>',
        unsafe_allow_html=True
    )
    st.markdown("""
<div class="info-box">
    🔵 <strong>Abu Dhabi:</strong>
    ADOSH-SF Version 4.0 (July 2024) |
    ADPHC | 54+ Codes of Practice |
    MOHRE Resolution 44/2022
</div>
<div class="info-box" style="margin-top:6px;">
    🔴 <strong>Dubai:</strong>
    Dubai Municipality Code of Construction
    Safety Practice | 23 Chapters |
    Local Order 61/1991
</div>
<div class="info-box" style="margin-top:6px;">
    🌍 <strong>International:</strong>
    ISO 45001:2018 | ISO 14001:2015 |
    NEBOSH | IOSH
</div>
""", unsafe_allow_html=True)

    st.divider()

    # ── SUMMER MIDDAY BAN ──────────────────────────────
    today = date.today()
    is_ban = (
        (today.month == 6 and today.day >= 15) or
        today.month == 7 or
        today.month == 8 or
        (today.month == 9 and today.day <= 15)
    )

    if is_ban:
        st.error(
            "☀️ **MIDDAY WORK BAN ACTIVE**\n\n"
            "No outdoor work: **12:30 — 15:00**\n\n"
            "Period: 15 June – 15 September\n\n"
            "Ref: MOHRE Resolution 44/2022 | "
            "ADOSH CoP 11.0"
        )
    else:
        st.info(
            "☀️ **Summer Midday Work Ban:**\n\n"
            "No outdoor work 12:30–15:00\n\n"
            "Active: 15 June – 15 September\n\n"
            "Ref: MOHRE Resolution 44/2022"
        )

st.divider()

# ═══════════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════════
st.markdown(f"""
<div style="
    background: #1F3864;
    border-radius: 8px;
    padding: 14px 20px;
    color: white;
    text-align: center;
    font-size: 12px;
    margin-top: 10px;
">
    <strong>UAE Construction HSE Compliance App</strong>
    &nbsp;|&nbsp;
    Developed by: Muhammad Aftab Qamar —
    Senior HSSE Engineer, ENGC
    &nbsp;|&nbsp;
    ADOSH-SF v4.0 | Dubai Municipality Code
    &nbsp;|&nbsp;
    © {date.today().year} ENGC — Bloom Living Projects,
    Abu Dhabi, UAE
    &nbsp;|&nbsp;
    <em>AI outputs must be verified by a competent
    HSE professional before operational use.</em>
</div>
""", unsafe_allow_html=True)