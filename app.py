import streamlit as st
from dotenv import load_dotenv
import os
from datetime import date
import hashlib
import time

load_dotenv()

# ── PAGE CONFIG — MUST BE FIRST ───────────────────────────
st.set_page_config(
    page_title="UAE HSE Compliance App — ENGC",
    page_icon="🦺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── YOUR DETAILS ──────────────────────────────────────────
YOUR_NAME        = "Muhammad Aftab Qamar"
YOUR_TITLE       = "Senior HSSE Engineer | ENGC"
YOUR_WHATSAPP    = "+971 52 267 1122"
YOUR_EMAIL       = "aftabqamar1122@hotmail.com"
FREE_TRIAL_MINS  = 15
MASTER_PWD       = os.getenv("MASTER_PASSWORD",
                              "ENGC@HSE2024")
SHARE_BASE       = os.getenv("SHARE_PASSWORD",
                              "ENGC_BLOOM_2024")

ADCB_BANK        = "Abu Dhabi Commercial Bank (ADCB)"
ADCB_TITLE       = "MUHAMMAD AFTAB QAMAR"
ADCB_ACC         = "11390882810001"
ADCB_IBAN        = "AE020030011390882810001"

ALFALAH_BANK     = "Bank Alfalah"
ALFALAH_TITLE    = "MUHAMMAD AFTAB QAMAR"
ALFALAH_ACC      = "56055002773488"
ALFALAH_IBAN     = "PK90ALFH5605005002773488"
ALFALAH_SWIFT    = "ALFHPKKAXXX"
ALFALAH_BRANCH   = "Ferozpur Road Lahore IBG"
ALFALAH_CODE     = "5605"

JAZZCASH_NO      = "+92 321 435 9532"
JAZZCASH_NAME    = "MUHAMMAD AFTAB QAMAR"


# ── AUTH HELPERS ──────────────────────────────────────────

def get_timed_pwd():
    now  = date.today()
    hr   = __import__('datetime').datetime.utcnow().hour
    ts   = f"{now.strftime('%Y%m%d')}{hr:02d}"
    return hashlib.sha256(
        f"{SHARE_BASE}{ts}".encode()
    ).hexdigest()[:8].upper()


def trial_expired():
    start = st.session_state.get("trial_start", 0)
    if not start:
        return False
    return (time.time() - start) / 60 >= FREE_TRIAL_MINS


def is_authenticated():
    if st.session_state.get("authenticated"):
        exp = st.session_state.get(
            "subscription_expiry", 0
        )
        if time.time() > exp:
            st.session_state.authenticated = False
            return False
        return True
    return False


def trial_running():
    return (
        st.session_state.get("trial_started")
        and not trial_expired()
    )


# ── CUSTOM CSS ────────────────────────────────────────────
st.markdown("""
<style>

/* Sidebar dark blue gradient */
[data-testid="stSidebar"] {
    background: linear-gradient(
        180deg,
        #1F3864 0%,
        #162847 50%,
        #0d1a33 100%
    ) !important;
    min-width: 255px !important;
    max-width: 255px !important;
}

/* All sidebar text white */
[data-testid="stSidebar"],
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div {
    color: rgba(255,255,255,0.9) !important;
}

/* Sidebar nav links */
[data-testid="stSidebarNav"] a {
    padding: 9px 16px !important;
    margin: 2px 6px !important;
    border-radius: 8px !important;
    color: rgba(255,255,255,0.8) !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    border-left: 3px solid transparent !important;
    display: flex !important;
    align-items: center !important;
}
[data-testid="stSidebarNav"] a:hover {
    background: rgba(255,255,255,0.1) !important;
    border-left: 3px solid #FFC000 !important;
    color: white !important;
}
[data-testid="stSidebarNav"] a[aria-current="page"] {
    background: rgba(255,255,255,0.15) !important;
    border-left: 3px solid #C00000 !important;
    color: white !important;
    font-weight: 700 !important;
}

/* Sidebar buttons */
[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,255,255,0.1) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.25) !important;
    border-radius: 8px !important;
    font-size: 12px !important;
    padding: 6px 4px !important;
    width: 100% !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.2) !important;
    border-color: #FFC000 !important;
    color: #FFC000 !important;
}

/* Hide Streamlit default elements */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* Main background */
.main {background-color: #f4f6f9 !important;}

/* Trial banner */
.trial-bar {
    background: linear-gradient(
        90deg, #1B5E20, #2E7D32
    );
    color: white;
    padding: 10px 20px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 600;
    margin-bottom: 16px;
    text-align: center;
}
.trial-warning {
    background: linear-gradient(
        90deg, #E65100, #BF360C
    );
    color: white;
    padding: 10px 20px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 600;
    margin-bottom: 16px;
    text-align: center;
}

/* Payment method box */
.pay-box {
    background: #F8F9FA;
    border: 2px solid #1F3864;
    border-radius: 12px;
    padding: 20px;
    margin: 8px 0;
}
.pay-row {
    display: flex;
    padding: 7px 0;
    border-bottom: 1px solid #E8E8E8;
    font-size: 13px;
}
.pay-lbl {
    font-weight: 700;
    color: #333;
    min-width: 160px;
}
.pay-val {
    color: #1F3864;
    font-weight: 600;
}

</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# PAYWALL PAGE
# ══════════════════════════════════════════════════════════

def show_paywall():
    st.markdown("""
<div style="
    background:linear-gradient(
        135deg,#1F3864,#C00000);
    border-radius:16px;padding:36px;
    color:white;text-align:center;
    margin-bottom:24px;">
    <div style="font-size:52px;">⏰</div>
    <div style="font-size:28px;
        font-weight:800;margin:10px 0;">
        Your Free Trial Has Ended
    </div>
    <div style="font-size:14px;opacity:0.9;">
        Subscribe to continue using the
        UAE HSE Compliance App
    </div>
</div>
""", unsafe_allow_html=True)

    c1, c2 = st.columns([1, 1.2])

    with c1:
        st.markdown("""
<div style="
    background:white;
    border:2px solid #1F3864;
    border-radius:12px;
    padding:24px;text-align:center;">
    <div style="color:#888;font-size:12px;
        letter-spacing:1px;">
        FULL ACCESS — 30 DAYS
    </div>
    <div style="font-size:60px;
        font-weight:800;color:#1F3864;
        line-height:1.1;">
        50
        <span style="font-size:22px;">AED</span>
    </div>
    <div style="color:#888;font-size:13px;
        margin-bottom:16px;">per month</div>
    <hr>
    <div style="text-align:left;
        font-size:13px;color:#333;
        line-height:1.9;">
        ✅ AI HSE Chatbot (Abu Dhabi + Dubai)<br>
        ✅ Risk Assessment Generator (.docx)<br>
        ✅ Toolbox Talk Generator<br>
        ✅ Permit to Work System<br>
        ✅ Incident Report Generator<br>
        ✅ Knowledge Base — 54 ADOSH CoPs<br>
        ✅ Knowledge Base — 23 Dubai Chapters<br>
        ✅ Unlimited document downloads
    </div>
</div>
""", unsafe_allow_html=True)

    with c2:
        st.markdown("### 📋 How to Subscribe")
        st.info(
            "**3 Easy Steps:**\n\n"
            "**1️⃣** Choose payment method below\n\n"
            "**2️⃣** Send **50 AED** with your "
            "Name + Email\n\n"
            "**3️⃣** Send payment screenshot "
            "via WhatsApp:\n\n"
            f"📱 **{YOUR_WHATSAPP}**\n\n"
            "✅ Access code sent within **1 hour**\n\n"
            "📅 **30 days** full access activated"
        )
        st.success(
            f"📱 WhatsApp: **{YOUR_WHATSAPP}**\n\n"
            f"✉️ Email: **{YOUR_EMAIL}**"
        )

    st.divider()
    st.markdown("## 💳 Payment Methods")

    t1, t2, t3 = st.tabs([
        "🏦 ADCB Bank (UAE — AED)",
        "🏦 Bank Alfalah (Pakistan)",
        "📱 JazzCash (Pakistan)",
    ])

    with t1:
        st.markdown(f"""
<div class="pay-box">
    <h3 style="color:#1F3864;">
        🏦 ADCB Bank — UAE (Pay in AED)
    </h3>
    <div class="pay-row">
        <span class="pay-lbl">Bank Name</span>
        <span class="pay-val">{ADCB_BANK}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Account Title</span>
        <span class="pay-val">{ADCB_TITLE}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Account Number</span>
        <span class="pay-val">{ADCB_ACC}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">IBAN</span>
        <span class="pay-val">{ADCB_IBAN}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Amount</span>
        <span class="pay-val"
            style="color:#C00000;
                   font-size:16px;">
            50 AED
        </span>
    </div>
    <div style="
        background:#E8F5E9;
        border-left:4px solid #2E7D32;
        border-radius:6px;
        padding:12px;margin-top:12px;
        font-size:13px;">
        ✅ After transfer, WhatsApp screenshot to:
        <strong>{YOUR_WHATSAPP}</strong><br>
        📝 Include: <strong>Full Name + Email</strong>
    </div>
</div>
""", unsafe_allow_html=True)

    with t2:
        st.markdown(f"""
<div class="pay-box">
    <h3 style="color:#1F3864;">
        🏦 Bank Alfalah — Pakistan (Pay in PKR)
    </h3>
    <div class="pay-row">
        <span class="pay-lbl">Bank Name</span>
        <span class="pay-val">{ALFALAH_BANK}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Account Title</span>
        <span class="pay-val">{ALFALAH_TITLE}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Account Number</span>
        <span class="pay-val">{ALFALAH_ACC}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">IBAN</span>
        <span class="pay-val">{ALFALAH_IBAN}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Swift Code</span>
        <span class="pay-val">{ALFALAH_SWIFT}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Branch</span>
        <span class="pay-val">{ALFALAH_BRANCH}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Branch Code</span>
        <span class="pay-val">{ALFALAH_CODE}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Amount</span>
        <span class="pay-val"
            style="color:#C00000;">
            PKR equivalent of 50 AED
            (~3,800–4,200 PKR)
        </span>
    </div>
    <div style="
        background:#E8F5E9;
        border-left:4px solid #2E7D32;
        border-radius:6px;
        padding:12px;margin-top:12px;
        font-size:13px;">
        ✅ After transfer, WhatsApp screenshot to:
        <strong>{YOUR_WHATSAPP}</strong><br>
        📝 Include: <strong>Full Name + Email</strong>
    </div>
</div>
""", unsafe_allow_html=True)

    with t3:
        st.markdown(f"""
<div class="pay-box">
    <h3 style="color:#1F3864;">
        📱 JazzCash — Pakistan (Pay in PKR)
    </h3>
    <div class="pay-row">
        <span class="pay-lbl">JazzCash Number</span>
        <span class="pay-val">{JAZZCASH_NO}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Account Name</span>
        <span class="pay-val">{JAZZCASH_NAME}</span>
    </div>
    <div class="pay-row">
        <span class="pay-lbl">Amount</span>
        <span class="pay-val"
            style="color:#C00000;">
            PKR equivalent of 50 AED
            (~3,800–4,200 PKR)
        </span>
    </div>
    <div style="
        background:#E8F5E9;
        border-left:4px solid #2E7D32;
        border-radius:6px;
        padding:12px;margin-top:12px;
        font-size:13px;">
        ✅ Open JazzCash → Send Money →
        Mobile Account<br>
        → Enter <strong>{JAZZCASH_NO}</strong>
        → Send amount<br>
        📸 WhatsApp screenshot to:
        <strong>{YOUR_WHATSAPP}</strong><br>
        📝 Include: <strong>Full Name + Email</strong>
    </div>
</div>
""", unsafe_allow_html=True)

    st.divider()
    st.markdown("### ✅ Already Paid? Enter Your Code")

    ca, cb = st.columns(2)
    with ca:
        code = st.text_input(
            "Access Code",
            type="password",
            placeholder="Enter code from WhatsApp",
            key="paywall_code"
        )
        if st.button(
            "🔓 Activate Subscription",
            type="primary",
            use_container_width=True,
            key="activate_btn"
        ):
            if code in [
                get_timed_pwd(), MASTER_PWD
            ]:
                st.session_state.authenticated = True
                st.session_state.auth_type = "sub"
                st.session_state.subscription_expiry\
                    = time.time() + 30*24*3600
                st.success(
                    "🎉 30 days access activated!"
                )
                time.sleep(1)
                st.rerun()
            else:
                st.error(
                    f"❌ Wrong code.\n"
                    f"Contact: {YOUR_WHATSAPP}"
                )
    with cb:
        st.info(
            "📱 Access code is sent via WhatsApp "
            "after payment verification.\n\n"
            f"Contact: **{YOUR_WHATSAPP}**\n\n"
            "Response within 1 hour"
        )

    if st.button("← Back to Login"):
        for k in ["trial_started", "trial_start",
                  "show_paywall"]:
            st.session_state[k] = False
        st.rerun()


# ══════════════════════════════════════════════════════════
# LOGIN PAGE
# ══════════════════════════════════════════════════════════

def show_login():
    st.markdown(f"""
<div style="
    background:linear-gradient(
        135deg,#1F3864 0%,#C00000 100%);
    border-radius:16px;padding:40px 32px;
    text-align:center;color:white;
    margin-bottom:30px;">
    <div style="font-size:52px;">🦺</div>
    <div style="font-size:26px;
        font-weight:800;margin:10px 0 6px 0;">
        UAE HSE Compliance App
    </div>
    <div style="font-size:14px;opacity:0.88;">
        Abu Dhabi &amp; Dubai &nbsp;|&nbsp;
        ADOSH-SF v4.0 &nbsp;|&nbsp;
        Dubai Municipality Code
    </div>
    <hr style="border-color:rgba(255,255,255,0.3);
        margin:14px 0;">
    <div style="font-size:13px;opacity:0.8;">
        Developed by: <strong>{YOUR_NAME}</strong>
        | {YOUR_TITLE}
    </div>
    <div style="font-size:12px;
        opacity:0.65;margin-top:4px;">
        📱 {YOUR_WHATSAPP} &nbsp;|&nbsp;
        ✉️ {YOUR_EMAIL}
    </div>
</div>
""", unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("""
<div style="
    background:#E8F5E9;
    border:2px solid #2E7D32;
    border-radius:12px;padding:22px;
    text-align:center;min-height:200px;">
    <div style="font-size:44px;">⏱️</div>
    <div style="font-size:18px;font-weight:700;
        color:#2E7D32;margin:8px 0;">
        15 Min Free Trial
    </div>
    <div style="font-size:12px;color:#555;
        line-height:1.6;">
        Try ALL features completely free
        for 15 minutes. No payment needed.
    </div>
</div>
""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button(
            "▶️ Start Free Trial",
            type="primary",
            use_container_width=True,
            key="start_trial"
        ):
            st.session_state.trial_started = True
            st.session_state.trial_start = time.time()
            st.rerun()

    with c2:
        st.markdown("""
<div style="
    background:#E8F0FE;
    border:2px solid #1F3864;
    border-radius:12px;padding:22px;
    text-align:center;min-height:200px;">
    <div style="font-size:44px;">🔑</div>
    <div style="font-size:18px;font-weight:700;
        color:#1F3864;margin:8px 0;">
        Enter Access Code
    </div>
    <div style="font-size:12px;color:#555;
        line-height:1.6;">
        Already subscribed?
        Enter your access code to login.
    </div>
</div>
""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        code = st.text_input(
            "Code",
            type="password",
            placeholder="Enter your code...",
            label_visibility="collapsed",
            key="login_code"
        )
        if st.button(
            "🔓 Login",
            use_container_width=True,
            key="login_btn"
        ):
            if code in [get_timed_pwd(), MASTER_PWD]:
                st.session_state.authenticated = True
                st.session_state.auth_type = "sub"
                st.session_state.subscription_expiry\
                    = time.time() + 30*24*3600
                st.success("✅ Welcome!")
                time.sleep(1)
                st.rerun()
            elif code:
                st.error(
                    f"❌ Wrong code.\n"
                    f"Contact: {YOUR_WHATSAPP}"
                )

    with c3:
        st.markdown(f"""
<div style="
    background:#FFF3E0;
    border:2px solid #E65100;
    border-radius:12px;padding:22px;
    text-align:center;min-height:200px;">
    <div style="font-size:44px;">💎</div>
    <div style="font-size:18px;font-weight:700;
        color:#E65100;margin:8px 0;">
        Subscribe Now
    </div>
    <div style="font-size:26px;font-weight:800;
        color:#E65100;">
        50 AED
        <span style="font-size:14px;
            font-weight:400;">/month</span>
    </div>
    <div style="font-size:12px;color:#555;">
        JazzCash · Alfalah · ADCB
    </div>
</div>
""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button(
            "💳 Subscribe — 50 AED/month",
            use_container_width=True,
            key="subscribe_btn"
        ):
            st.session_state.show_paywall = True
            st.rerun()

    st.divider()
    st.markdown(f"""
<div style="
    background:#1F3864;border-radius:10px;
    padding:12px 20px;text-align:center;
    color:white;font-size:13px;">
    📞 <strong>WhatsApp:</strong>
    {YOUR_WHATSAPP} &nbsp;|&nbsp;
    ✉️ <strong>Email:</strong> {YOUR_EMAIL}
    &nbsp;|&nbsp;
    ⏰ Response within 1 hour
    (Abu Dhabi business hours)
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# ACCESS CONTROL GATE
# ══════════════════════════════════════════════════════════

# Show paywall if requested
if st.session_state.get("show_paywall"):
    show_paywall()
    st.stop()

# Show paywall if trial expired
if (st.session_state.get("trial_started")
        and trial_expired()):
    show_paywall()
    st.stop()

# Show login if not authenticated and no trial
if not is_authenticated() and not trial_running():
    show_login()
    st.stop()

# Show trial countdown banner
if trial_running():
    start   = st.session_state.get("trial_start", 0)
    elapsed = (time.time() - start) / 60
    mins    = max(0, FREE_TRIAL_MINS - elapsed)
    mi      = int(mins)
    se      = int((mins - mi) * 60)
    if mins > 5:
        st.markdown(f"""
<div class="trial-bar">
    ⏱️ FREE TRIAL: <strong>{mi}m {se}s
    remaining</strong> — Subscribe for full
    access: <strong>50 AED / month</strong> |
    WhatsApp: {YOUR_WHATSAPP}
</div>
""", unsafe_allow_html=True)
    else:
        st.markdown(f"""
<div class="trial-warning">
    ⚠️ TRIAL ENDING SOON:
    <strong>{mi}m {se}s left!</strong>
    Subscribe NOW → WhatsApp:
    {YOUR_WHATSAPP}
    &nbsp;&nbsp;
</div>
""", unsafe_allow_html=True)
        if st.button(
            "💳 Subscribe Now — 50 AED/month",
            key="subscribe_now_banner"
        ):
            st.session_state.show_paywall = True
            st.rerun()


# ══════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════

with st.sidebar:

    st.markdown("""
<div style="text-align:center;
    padding:16px 8px 8px 8px;">
    <div style="font-size:38px;">🦺</div>
    <div style="font-size:15px;font-weight:800;
        color:white;letter-spacing:0.5px;
        margin-top:6px;line-height:1.3;">
        UAE HSE<br>Compliance App
    </div>
    <div style="font-size:10px;
        color:rgba(255,255,255,0.55);
        margin-top:4px;">
        ADOSH-SF v4.0 | DM Code
    </div>
</div>
<hr style="border-color:rgba(255,255,255,0.15);
    margin:8px 0;">
""", unsafe_allow_html=True)

    # Emirate selector in sidebar
    st.markdown("""
<div style="font-size:10px;font-weight:700;
    color:rgba(255,255,255,0.55);
    letter-spacing:1.5px;
    text-transform:uppercase;
    padding:0 4px;margin-bottom:6px;">
    🗺️ Select Emirate
</div>
""", unsafe_allow_html=True)

    if "emirate" not in st.session_state:
        st.session_state.emirate = None

    sb1, sb2 = st.columns(2)
    with sb1:
        if st.button(
            "🔵 Abu\nDhabi",
            key="sb_ad",
            use_container_width=True
        ):
            st.session_state.emirate = "Abu Dhabi"
            st.rerun()
    with sb2:
        if st.button(
            "🔴\nDubai",
            key="sb_dxb",
            use_container_width=True
        ):
            st.session_state.emirate = "Dubai"
            st.rerun()

    if st.session_state.emirate == "Abu Dhabi":
        st.markdown("""
<div style="background:rgba(0,100,200,0.2);
    border:1px solid rgba(100,180,255,0.35);
    border-radius:8px;padding:8px 10px;
    font-size:11px;color:#90CAF9;
    text-align:center;margin:6px 0;">
    ✅ <strong>Abu Dhabi</strong> Active<br>
    <span style="font-size:10px;
        color:rgba(255,255,255,0.45);">
        ADOSH-SF v4.0 | ADPHC
    </span>
</div>
""", unsafe_allow_html=True)
    elif st.session_state.emirate == "Dubai":
        st.markdown("""
<div style="background:rgba(200,0,0,0.2);
    border:1px solid rgba(255,100,100,0.35);
    border-radius:8px;padding:8px 10px;
    font-size:11px;color:#EF9A9A;
    text-align:center;margin:6px 0;">
    ✅ <strong>Dubai</strong> Active<br>
    <span style="font-size:10px;
        color:rgba(255,255,255,0.45);">
        Dubai Municipality Code
    </span>
</div>
""", unsafe_allow_html=True)
    else:
        st.markdown("""
<div style="background:rgba(255,200,0,0.12);
    border:1px solid rgba(255,200,0,0.3);
    border-radius:8px;padding:8px 10px;
    font-size:11px;color:#FFC000;
    text-align:center;margin:6px 0;">
    ⬆️ Select an Emirate above
</div>
""", unsafe_allow_html=True)

    st.markdown("""
<hr style="border-color:rgba(255,255,255,0.12);
    margin:8px 0;">
<div style="font-size:10px;font-weight:700;
    color:rgba(255,255,255,0.5);
    letter-spacing:1.5px;
    text-transform:uppercase;
    padding:0 4px;margin-bottom:6px;">
    📱 App Modules
</div>
""", unsafe_allow_html=True)

    mods = [
        ("📚", "Knowledge Base"),
        ("🤖", "AI Chatbot"),
        ("📋", "Risk Assessment"),
        ("🗣️", "Toolbox Talks"),
        ("📝", "Permit to Work"),
        ("🚨", "Incident Report"),
    ]
    for icon, name in mods:
        st.markdown(f"""
<div style="display:flex;align-items:center;
    padding:8px 12px;margin:3px 0;
    border-radius:8px;
    background:rgba(255,255,255,0.06);
    border-left:3px solid
    rgba(255,255,255,0.12);
    font-size:13px;
    color:rgba(255,255,255,0.85);">
    <span style="font-size:15px;
        margin-right:10px;">{icon}</span>
    {name}
</div>
""", unsafe_allow_html=True)

    st.markdown("""
<hr style="border-color:rgba(255,255,255,0.12);
    margin:8px 0;">
<div style="font-size:10px;font-weight:700;
    color:rgba(255,255,255,0.5);
    letter-spacing:1.5px;
    text-transform:uppercase;
    padding:0 4px;margin-bottom:6px;">
    🚨 Emergency Numbers
</div>
<div style="background:rgba(192,0,0,0.18);
    border:1px solid rgba(255,100,100,0.25);
    border-radius:8px;padding:10px 12px;
    font-size:12px;line-height:2;
    color:rgba(255,255,255,0.88);">
    🚔 Police: <strong>999</strong><br>
    🚑 Ambulance: <strong>998</strong><br>
    🚒 Civil Defence: <strong>997</strong><br>
    🏛️ DM Emergency: <strong>800900</strong>
</div>
""", unsafe_allow_html=True)

    today = date.today()
    is_ban = (
        (today.month == 6 and today.day >= 15)
        or today.month in [7, 8]
        or (today.month == 9 and today.day <= 15)
    )
    if is_ban:
        st.markdown("""
<div style="background:rgba(255,100,0,0.22);
    border:1px solid rgba(255,150,0,0.4);
    border-radius:8px;padding:10px 12px;
    font-size:11px;color:#FFB74D;
    text-align:center;
    line-height:1.7;margin-top:8px;">
    ☀️ <strong>MIDDAY BAN ACTIVE</strong><br>
    No outdoor work<br>
    <strong>12:30 — 15:00</strong><br>
    15 Jun – 15 Sep<br>
    <span style="font-size:10px;
        color:rgba(255,255,255,0.45);">
        MOHRE Res. 44/2022
    </span>
</div>
""", unsafe_allow_html=True)

    st.markdown(f"""
<hr style="border-color:rgba(255,255,255,0.12);
    margin:8px 0;">
<div style="text-align:center;padding:4px;
    font-size:10px;
    color:rgba(255,255,255,0.4);
    line-height:1.7;">
    Developed by<br>
    <strong style="
        color:rgba(255,255,255,0.65);">
        {YOUR_NAME}
    </strong><br>
    Senior HSSE Engineer | ENGC<br>
    📱 {YOUR_WHATSAPP}
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# MAIN PAGE CONTENT
# ══════════════════════════════════════════════════════════

st.markdown("""
<div style="
    background:linear-gradient(
        135deg,#1F3864 0%,#C00000 100%);
    border-radius:12px;
    padding:28px 32px;
    color:white;margin-bottom:20px;">
    <div style="font-size:30px;
        font-weight:800;margin:0;
        letter-spacing:1px;">
        🦺 UAE Construction HSE Compliance App
    </div>
    <div style="font-size:14px;
        opacity:0.9;margin-top:6px;">
        Abu Dhabi & Dubai | ADOSH-SF v4.0 |
        Dubai Municipality Code |
        AI-Powered HSE Management System
    </div>
</div>
""", unsafe_allow_html=True)

col_contact, col_selector = st.columns([1, 1.4])

with col_contact:
    st.markdown(f"""
<div style="
    background:linear-gradient(
        135deg,#1F3864 0%,#2E4F8A 100%);
    border-radius:12px;padding:20px;
    color:white;text-align:center;">
    <div style="font-size:19px;
        font-weight:700;margin-bottom:4px;">
        {YOUR_NAME}
    </div>
    <div style="font-size:12px;opacity:0.82;
        margin-bottom:12px;">
        Senior HSSE Engineer<br>
        Exeed National General Contracting
        LLC (ENGC)
    </div>
    <hr style="border-color:
        rgba(255,255,255,0.25);margin:10px 0;">
    <div style="font-size:13px;margin:4px 0;">
        📱 {YOUR_WHATSAPP}
    </div>
    <div style="font-size:13px;margin:4px 0;">
        ✉️ {YOUR_EMAIL}
    </div>
    <div style="font-size:13px;margin:4px 0;">
        📍 Abu Dhabi, UAE
    </div>
    <hr style="border-color:
        rgba(255,255,255,0.25);margin:10px 0;">
    <div style="font-size:10px;opacity:0.7;">
        15 Yrs+ Experience | NEBOSH Idip |
        ADOSH Senior Practitioner |
        ISO 45001 | ADOSH-SF v4.0
    </div>
</div>
""", unsafe_allow_html=True)

with col_selector:
    st.markdown("""
<div style="background:#1F3864;color:white;
    padding:8px 14px;border-radius:6px;
    font-weight:700;font-size:13px;
    margin-bottom:10px;">
    🗺️ SELECT YOUR EMIRATE TO BEGIN
</div>
""", unsafe_allow_html=True)

    if "emirate" not in st.session_state:
        st.session_state.emirate = None

    mc1, mc2 = st.columns(2)
    with mc1:
        if st.button(
            "🔵 ABU DHABI\nADOSH-SF v4.0",
            use_container_width=True,
            type="primary",
            key="main_ad"
        ):
            st.session_state.emirate = "Abu Dhabi"
            st.rerun()
    with mc2:
        if st.button(
            "🔴 DUBAI\nDM Code",
            use_container_width=True,
            type="secondary",
            key="main_dxb"
        ):
            st.session_state.emirate = "Dubai"
            st.rerun()

    if st.session_state.emirate == "Abu Dhabi":
        st.success(
            "✅ **Abu Dhabi** selected — "
            "ADOSH-SF v4.0 framework active"
        )
        st.info("👈 Use the sidebar to navigate")
    elif st.session_state.emirate == "Dubai":
        st.error(
            "✅ **Dubai** selected — "
            "Dubai Municipality Code active"
        )
        st.info("👈 Use the sidebar to navigate")
    else:
        st.warning(
            "⬆️ Please select your Emirate "
            "above to activate the app"
        )

    st.divider()
    st.markdown("""
<div style="background:#1F3864;color:white;
    padding:8px 14px;border-radius:6px;
    font-weight:700;font-size:13px;
    margin-bottom:8px;">
    🚨 UAE EMERGENCY NUMBERS
</div>
""", unsafe_allow_html=True)
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Police", "999")
    e2.metric("Ambulance", "998")
    e3.metric("Civil Defence", "997")
    e4.metric("DM Emergency", "800900")

st.divider()

# Projects + Modules
col_proj, col_mods = st.columns([1.3, 1])

with col_proj:
    st.markdown("""
<div style="background:#1F3864;color:white;
    padding:8px 14px;border-radius:6px;
    font-weight:700;font-size:13px;
    margin-bottom:10px;">
    🏗️ BLOOM LIVING PROJECTS — ABU DHABI
</div>
""", unsafe_allow_html=True)

    pi1, pi2, pi3 = st.columns(3)
    for col, label, val in [
        (pi1, "CLIENT",
         "Bloom Living / Bloom District"),
        (pi2, "PMC",
         "TIMES"),
        (pi3, "CONSULTANT",
         "Dar Al-Handasah"),
    ]:
        col.markdown(f"""
<div style="background:#EEF2FF;
    border-radius:8px;padding:10px;
    font-size:12px;color:#1F3864;
    margin:3px 0;">
    <strong>{label}</strong><br>{val}
</div>
""", unsafe_allow_html=True)

    pi4, pi5 = st.columns(2)
    pi4.markdown("""
<div style="background:#EEF2FF;
    border-radius:8px;padding:10px;
    font-size:12px;color:#1F3864;
    margin:3px 0;">
    <strong>MAIN CONTRACTOR</strong><br>
    ENGC
</div>
""", unsafe_allow_html=True)
    pi5.markdown("""
<div style="background:#EEF2FF;
    border-radius:8px;padding:10px;
    font-size:12px;color:#1F3864;
    margin:3px 0;">
    <strong>SUB-CONTRACTOR</strong><br>
    ___________________________
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Villa Projects
    st.markdown("""
<div style="background:#8B0000;color:white;
    padding:7px 14px;border-radius:6px;
    font-weight:700;font-size:12px;
    margin-bottom:6px;">
    🏡 RESIDENTIAL VILLA PROJECTS
</div>
""", unsafe_allow_html=True)

    for name, ref in [
        ("Olvera — Bloom Living",  "BL-VL-01"),
        ("Almeria — Bloom Living", "BL-VL-02"),
    ]:
        st.markdown(f"""
<div style="background:white;
    border-left:5px solid #8B0000;
    border-radius:8px;padding:10px 14px;
    margin:4px 0;
    box-shadow:0 1px 4px rgba(0,0,0,0.08);">
    <div style="font-weight:700;
        font-size:13px;color:#8B0000;">
        🏡 {name}
    </div>
    <div style="font-size:11px;color:#666;">
        Residential Villa | Ref: {ref} |
        Abu Dhabi, UAE
    </div>
</div>
""", unsafe_allow_html=True)

    # Community Centers
    st.markdown("""
<div style="background:#1F3864;color:white;
    padding:7px 14px;border-radius:6px;
    font-weight:700;font-size:12px;
    margin:10px 0 6px 0;">
    🏛️ COMMUNITY CENTER PROJECTS
</div>
""", unsafe_allow_html=True)

    for name, ref in [
        ("Cordoba Community Center", "BL-CC-01"),
        ("Toledo Community Center",  "BL-CC-02"),
        ("Casares Community Center", "BL-CC-03"),
    ]:
        st.markdown(f"""
<div style="background:white;
    border-left:5px solid #1F3864;
    border-radius:8px;padding:10px 14px;
    margin:4px 0;
    box-shadow:0 1px 4px rgba(0,0,0,0.08);">
    <div style="font-weight:700;
        font-size:13px;color:#1F3864;">
        🏛️ {name}
    </div>
    <div style="font-size:11px;color:#666;">
        Community Center | Ref: {ref} |
        Abu Dhabi, UAE
    </div>
</div>
""", unsafe_allow_html=True)

    # Service Station
    st.markdown("""
<div style="background:#5C4033;color:white;
    padding:7px 14px;border-radius:6px;
    font-weight:700;font-size:12px;
    margin:10px 0 6px 0;">
    ⛽ SERVICE STATION PROJECT
</div>
""", unsafe_allow_html=True)
    st.markdown("""
<div style="background:white;
    border-left:5px solid #5C4033;
    border-radius:8px;padding:10px 14px;
    margin:4px 0;
    box-shadow:0 1px 4px rgba(0,0,0,0.08);">
    <div style="font-weight:700;
        font-size:13px;color:#5C4033;">
        ⛽ Service Station — Bloom Living
    </div>
    <div style="font-size:11px;color:#666;">
        Service Center | Plot C13, Zayed City |
        Ref: BL-SS-01 | Abu Dhabi, UAE
    </div>
</div>
""", unsafe_allow_html=True)

with col_mods:
    st.markdown("""
<div style="background:#1F3864;color:white;
    padding:8px 14px;border-radius:6px;
    font-weight:700;font-size:13px;
    margin-bottom:10px;">
    📱 APP MODULES
</div>
""", unsafe_allow_html=True)

    mods_main = [
        ("📚", "Knowledge Base",
         "#1F3864",
         "Browse 54 ADOSH CoPs or 23 Dubai "
         "Code Chapters with AI search"),
        ("🤖", "AI HSE Chatbot",
         "#0D47A1",
         "Ask any HSE question — Emirates-specific "
         "answers with CoP references"),
        ("📋", "Risk Assessment",
         "#1B5E20",
         "ENGC 14-column HIRA with hierarchy "
         "of controls — downloads as .docx"),
        ("🗣️", "Toolbox Talks",
         "#4A148C",
         "AI-powered toolbox talks for any HSE "
         "topic with attendance sheet"),
        ("📝", "Permit to Work",
         "#BF360C",
         "Digital PTW — Hot Work, Confined Space, "
         "Excavation, Heights, LOTO, Lifting"),
        ("🚨", "Incident Report",
         "#C00000",
         "Full investigation with 5-Why analysis "
         "and corrective action tables"),
    ]
    for icon, name, color, desc in mods_main:
        st.markdown(f"""
<div style="background:white;
    border-left:5px solid {color};
    border-radius:8px;padding:10px 14px;
    margin:5px 0;
    box-shadow:0 1px 4px rgba(0,0,0,0.07);">
    <div style="font-weight:700;
        font-size:14px;color:{color};">
        {icon} {name}
    </div>
    <div style="font-size:11px;
        color:#666;margin-top:2px;">
        {desc}
    </div>
</div>
""", unsafe_allow_html=True)

    st.divider()

    today = date.today()
    is_ban = (
        (today.month == 6 and today.day >= 15)
        or today.month in [7, 8]
        or (today.month == 9 and today.day <= 15)
    )
    if is_ban:
        st.error(
            "☀️ **MIDDAY WORK BAN ACTIVE**\n\n"
            "No outdoor work: **12:30 — 15:00**\n\n"
            "15 June – 15 September\n\n"
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

st.markdown(f"""
<div style="background:#1F3864;
    border-radius:8px;
    padding:14px 20px;
    color:white;text-align:center;
    font-size:12px;margin-top:10px;">
    <strong>UAE Construction HSE Compliance App
    </strong> &nbsp;|&nbsp;
    Developed by: {YOUR_NAME} —
    Senior HSSE Engineer, ENGC
    &nbsp;|&nbsp;
    ADOSH-SF v4.0 | Dubai Municipality Code
    &nbsp;|&nbsp;
    © {date.today().year} ENGC — Bloom Living
    Projects, Abu Dhabi, UAE
    &nbsp;|&nbsp;
    <em>AI outputs must be verified by a
    competent HSE professional.</em>
</div>
""", unsafe_allow_html=True)