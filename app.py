import streamlit as st
import streamlit.components.v1
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

def ACCESS CONTROL GATE():
    # ══════════════════════════════════════════════════════════
# BROWSER STORAGE — Permanent Trial Tracking
# ══════════════════════════════════════════════════════════

def inject_trial_checker():
    """
    Inject JavaScript that:
    1. Checks if trial was already used
       (stored in browser localStorage)
    2. If used → sets URL parameter ?trial=used
    3. If not used → allows trial to start
    """
    st.components.v1.html("""
<script>
// Check localStorage for trial status
var trialUsed = localStorage.getItem(
    'hse_app_trial_used'
);
var trialExpiry = localStorage.getItem(
    'hse_app_trial_expiry'
);

// If trial was used and is expired
if (trialUsed === 'true') {
    // Tell Streamlit via URL param
    var url = new URL(window.location.href);
    if (!url.searchParams.get('trial_status')) {
        url.searchParams.set(
            'trial_status', 'expired'
        );
        window.location.href = url.toString();
    }
}
</script>
""", height=0)


def mark_trial_used_in_browser():
    """
    Inject JavaScript to permanently mark
    trial as used in this browser.
    """
    st.components.v1.html("""
<script>
// Save to localStorage — persists after
// browser close
localStorage.setItem(
    'hse_app_trial_used', 'true'
);
localStorage.setItem(
    'hse_app_trial_expiry',
    new Date().toISOString()
);
console.log('HSE App: Trial marked as used');
</script>
""", height=0)


def clear_trial_for_subscriber():
    """
    When someone subscribes, clear the trial
    restriction from their browser.
    """
    st.components.v1.html("""
<script>
localStorage.setItem(
    'hse_app_trial_used', 'subscribed'
);
localStorage.setItem(
    'hse_app_subscribed', 'true'
);
console.log('HSE App: Subscription activated');
</script>
""", height=0)


# ══════════════════════════════════════════════════════════
# LOGIN PAGE
# ══════════════════════════════════════════════════════════

def show_login(trial_blocked=False):
    """Main entry/login page."""

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
    <hr style="border-color:
        rgba(255,255,255,0.3);margin:14px 0;">
    <div style="font-size:13px;opacity:0.8;">
        Developed by:
        <strong>{YOUR_NAME}</strong>
        | {YOUR_TITLE}
    </div>
    <div style="font-size:12px;
        opacity:0.65;margin-top:4px;">
        📱 {YOUR_WHATSAPP} &nbsp;|&nbsp;
        ✉️ {YOUR_EMAIL}
    </div>
</div>
""", unsafe_allow_html=True)

    # ── Trial Already Used Warning ──────────────────────
    if trial_blocked:
        st.error(
            "⏰ **Your free trial has already been used.**\n\n"
            "The 15-minute free trial can only be used "
            "**once per device**.\n\n"
            "Please subscribe to continue using the app."
        )
        st.markdown("<br>", unsafe_allow_html=True)

    # ── Show 3 columns or 2 if trial blocked ───────────
    if trial_blocked:
        # Only show Login + Subscribe (no trial option)
        c1, c2 = st.columns(2)
        cols = [c1, c2]
        show_trial_col = False
    else:
        c1, c2, c3 = st.columns(3)
        cols = [c1, c2, c3]
        show_trial_col = True

    # Free Trial Column (only if not blocked)
    if show_trial_col:
        with c1:
            st.markdown("""
<div style="
    background:#E8F5E9;
    border:2px solid #2E7D32;
    border-radius:12px;padding:22px;
    text-align:center;min-height:200px;">
    <div style="font-size:44px;">⏱️</div>
    <div style="font-size:18px;
        font-weight:700;color:#2E7D32;
        margin:8px 0;">
        15 Min Free Trial
    </div>
    <div style="font-size:12px;
        color:#555;line-height:1.6;">
        Try ALL features free for 15 mins.
        <br><strong>One time only.</strong>
        No payment needed.
    </div>
</div>
""", unsafe_allow_html=True)
            st.markdown(
                "<br>", unsafe_allow_html=True
            )
            if st.button(
                "▶️ Start Free Trial",
                type="primary",
                use_container_width=True,
                key="start_trial"
            ):
                # Mark trial as started
                st.session_state.trial_started = True
                st.session_state.trial_start = (
                    time.time()
                )
                st.session_state.trial_used = True
                # Mark in browser storage
                mark_trial_used_in_browser()
                st.rerun()

    # Access Code Column
    code_col = c2 if show_trial_col else c1
    with code_col:
        st.markdown("""
<div style="
    background:#E8F0FE;
    border:2px solid #1F3864;
    border-radius:12px;padding:22px;
    text-align:center;min-height:200px;">
    <div style="font-size:44px;">🔑</div>
    <div style="font-size:18px;
        font-weight:700;color:#1F3864;
        margin:8px 0;">
        Enter Access Code
    </div>
    <div style="font-size:12px;
        color:#555;line-height:1.6;">
        Already subscribed?
        Enter your access code to login.
    </div>
</div>
""", unsafe_allow_html=True)
        st.markdown(
            "<br>", unsafe_allow_html=True
        )
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
            if code in [
                get_timed_pwd(), MASTER_PWD
            ]:
                st.session_state.authenticated = True
                st.session_state.auth_type = "sub"
                st.session_state.\
                    subscription_expiry = (
                    time.time() + 30*24*3600
                )
                # Clear trial restriction
                # for subscribers
                clear_trial_for_subscriber()
                st.success(
                    "✅ Welcome! Loading..."
                )
                time.sleep(1)
                st.rerun()
            elif code:
                st.error(
                    f"❌ Wrong code.\n"
                    f"Contact: {YOUR_WHATSAPP}"
                )

    # Subscribe Column
    sub_col = c3 if show_trial_col else c2
    with sub_col:
        st.markdown(f"""
<div style="
    background:#FFF3E0;
    border:2px solid #E65100;
    border-radius:12px;padding:22px;
    text-align:center;min-height:200px;">
    <div style="font-size:44px;">💎</div>
    <div style="font-size:18px;
        font-weight:700;color:#E65100;
        margin:8px 0;">
        Subscribe Now
    </div>
    <div style="font-size:26px;
        font-weight:800;color:#E65100;">
        50 AED
        <span style="font-size:14px;
            font-weight:400;">
            /month
        </span>
    </div>
    <div style="font-size:12px;color:#555;">
        JazzCash · Alfalah · ADCB
    </div>
</div>
""", unsafe_allow_html=True)
        st.markdown(
            "<br>", unsafe_allow_html=True
        )
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
    ✉️ <strong>Email:</strong>
    {YOUR_EMAIL} &nbsp;|&nbsp;
    ⏰ Response within 1 hour
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# ACCESS CONTROL GATE — WITH PERMANENT TRIAL TRACKING
# ══════════════════════════════════════════════════════════

# Check URL parameters for trial status
# (set by our JavaScript)
trial_status_param = st.query_params.get(
    "trial_status", ""
)
trial_blocked_by_browser = (
    trial_status_param == "expired"
)

# Show paywall if requested
if st.session_state.get("show_paywall"):
    show_paywall()
    st.stop()

# Show paywall if trial expired THIS session
if (st.session_state.get("trial_started")
        and trial_expired()):
    # Mark permanently in browser
    mark_trial_used_in_browser()
    show_paywall()
    st.stop()

# Check if authenticated (paid subscriber)
if is_authenticated():
    pass  # Continue to app

# Check if trial is currently running
elif trial_running():
    # Trial is active — show countdown
    start   = st.session_state.get(
        "trial_start", 0
    )
    elapsed = (time.time() - start) / 60
    mins    = max(0, FREE_TRIAL_MINS - elapsed)
    mi      = int(mins)
    se      = int((mins - mi) * 60)

    if mins > 5:
        st.markdown(f"""
<div class="trial-bar">
    ⏱️ FREE TRIAL: <strong>{mi}m {se}s
    remaining</strong> — One time only!
    Subscribe: <strong>50 AED / month</strong>
    | WhatsApp: {YOUR_WHATSAPP}
</div>
""", unsafe_allow_html=True)
    else:
        st.markdown(f"""
<div class="trial-warning">
    ⚠️ TRIAL ENDING: <strong>{mi}m {se}s
    left!</strong> Subscribe NOW →
    WhatsApp: {YOUR_WHATSAPP}
</div>
""", unsafe_allow_html=True)
        if st.button(
            "💳 Subscribe — 50 AED/month",
            key="sub_banner"
        ):
            st.session_state.show_paywall = True
            st.rerun()

# Trial already used (from browser storage)
elif trial_blocked_by_browser:
    show_login(trial_blocked=True)
    st.stop()

# Trial already marked as used this session
elif st.session_state.get("trial_used"):
    show_login(trial_blocked=True)
    st.stop()

# No access — show login
else:
    # Inject the browser checker first
    inject_trial_checker()
    show_login(trial_blocked=False)
    st.stop()