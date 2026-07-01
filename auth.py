import streamlit as st
import hashlib
import time
from datetime import datetime, timedelta
import os

# ── YOUR PERSONAL & PAYMENT DETAILS ──────────────────────
MASTER_PASSWORD     = os.getenv(
    "MASTER_PASSWORD", "ENGC@HSE2024"
)
SHARE_PASSWORD_BASE = os.getenv(
    "SHARE_PASSWORD", "ENGC_BLOOM_2024"
)
FREE_TRIAL_MINUTES  = 15

YOUR_NAME           = "Muhammad Aftab Qamar"
YOUR_TITLE          = "Senior HSSE Engineer | ENGC"
YOUR_WHATSAPP       = "+971 52 267 1122"
YOUR_EMAIL          = "aftabqamar1122@hotmail.com"

# ADCB Bank — UAE
ADCB_BANK_NAME      = (
    "Abu Dhabi Commercial Bank PJSC (ADCB)"
)
ADCB_ACCOUNT_TITLE  = "MUHAMMAD AFTAB QAMAR"
ADCB_ACCOUNT_NO     = "11390882810001"
ADCB_IBAN           = "AE020030011390882810001"

# Bank Alfalah — Pakistan
ALFALAH_BANK_NAME   = "Bank Alfalah"
ALFALAH_TITLE       = "MUHAMMAD AFTAB QAMAR"
ALFALAH_ACCOUNT_NO  = "56055002773488"
ALFALAH_IBAN        = "PK90ALFH5605005002773488"

# JazzCash — Pakistan
JAZZCASH_NUMBER     = "+92 321 435 9532"
JAZZCASH_NAME       = "MUHAMMAD AFTAB QAMAR"


# ── HELPERS ───────────────────────────────────────────────

def get_timed_password():
    """Password valid for current 1-hour window."""
    now   = datetime.utcnow()
    ad    = now + timedelta(hours=4)
    win   = ad.replace(
        minute=0, second=0, microsecond=0
    )
    ts    = win.strftime("%Y%m%d%H%M")
    combo = f"{SHARE_PASSWORD_BASE}{ts}"
    return hashlib.sha256(
        combo.encode()
    ).hexdigest()[:8].upper()


def trial_expired():
    """Check if 15-min free trial is over."""
    start = st.session_state.get("trial_start", 0)
    if not start:
        return False
    elapsed = (time.time() - start) / 60
    return elapsed >= FREE_TRIAL_MINUTES


def subscription_active():
    """Check if paid subscription is active."""
    sub_exp = st.session_state.get(
        "subscription_expiry", 0
    )
    if not sub_exp:
        return False
    return time.time() < sub_exp


def check_access():
    """
    Main access gate.
    Returns True if user can use the app.
    """
    if st.session_state.get("authenticated"):
        if not subscription_active():
            if st.session_state.get(
                "auth_type"
            ) == "subscription":
                st.session_state.authenticated = False
                return False
        return True
    if st.session_state.get("trial_started"):
        if trial_expired():
            return False
        return True
    return False


def show_trial_banner():
    """Show countdown banner during free trial."""
    start    = st.session_state.get("trial_start", 0)
    elapsed  = (time.time() - start) / 60
    mins     = max(0, FREE_TRIAL_MINUTES - elapsed)
    mins_int = int(mins)
    secs     = int((mins - mins_int) * 60)

    if mins > 5:
        st.success(
            f"⏱️ FREE TRIAL: "
            f"**{mins_int} min {secs} sec remaining**"
            f" — Subscribe for full access: "
            f"**50 AED / month**"
        )
    else:
        st.warning(
            f"⚠️ TRIAL ENDING SOON: "
            f"**{mins_int} min {secs} sec left!** "
            f"Subscribe NOW to keep access! "
            f"WhatsApp: **{YOUR_WHATSAPP}**"
        )


def show_paywall():
    """Full payment page shown when trial expires."""

    st.markdown("""
<style>
.pay-header {
    background: linear-gradient(
        135deg, #1F3864 0%, #C00000 100%
    );
    border-radius: 16px;
    padding: 36px 32px;
    color: white;
    text-align: center;
    margin-bottom: 24px;
}
.price-box {
    background: white;
    border: 2px solid #1F3864;
    border-radius: 12px;
    padding: 24px;
    text-align: center;
    height: 100%;
}
.pay-method {
    background: #F8F9FA;
    border: 2px solid #1F3864;
    border-radius: 12px;
    padding: 20px;
    margin: 8px 0;
}
.pay-method h3 {
    color: #1F3864;
    margin-bottom: 12px;
}
.pay-row {
    display: flex;
    justify-content: space-between;
    padding: 6px 0;
    border-bottom: 1px solid #E0E0E0;
    font-size: 13px;
}
.pay-label {
    font-weight: 700;
    color: #333;
    min-width: 140px;
}
.pay-value {
    color: #1F3864;
    font-weight: 600;
}
.important-note {
    background: #FFF3E0;
    border-left: 4px solid #E65100;
    border-radius: 6px;
    padding: 12px 16px;
    font-size: 13px;
    margin-top: 10px;
}
</style>
""", unsafe_allow_html=True)

    # ── HEADER ────────────────────────────────────────
    st.markdown(f"""
<div class="pay-header">
    <div style="font-size:52px;">⏰</div>
    <div style="font-size:28px;
                font-weight:800;
                margin:10px 0;">
        Your Free Trial Has Ended
    </div>
    <div style="font-size:15px; opacity:0.9;">
        Subscribe to continue using the
        UAE HSE Compliance App
    </div>
    <div style="font-size:12px;
                opacity:0.75;
                margin-top:8px;">
        Developed by: {YOUR_NAME} |
        {YOUR_TITLE}
    </div>
</div>
""", unsafe_allow_html=True)

    # ── PRICE + FEATURES ──────────────────────────────
    col1, col2 = st.columns([1, 1.2])

    with col1:
        st.markdown("""
<div class="price-box">
    <div style="color:#888;
                font-size:12px;
                letter-spacing:1px;">
        FULL ACCESS — 30 DAYS
    </div>
    <div style="font-size:58px;
                font-weight:800;
                color:#1F3864;
                line-height:1.1;">
        50
        <span style="font-size:22px;">AED</span>
    </div>
    <div style="color:#888;
                font-size:13px;
                margin-bottom:16px;">
        per month
    </div>
    <hr style="margin:12px 0;">
    <div style="text-align:left;
                font-size:13px;
                color:#333;
                line-height:1.8;">
        ✅ AI HSE Chatbot
        (Abu Dhabi + Dubai)<br>
        ✅ Risk Assessment Generator
        (.docx)<br>
        ✅ Toolbox Talk Generator<br>
        ✅ Permit to Work System<br>
        ✅ Incident Report Generator<br>
        ✅ Knowledge Base —
        54 ADOSH CoPs<br>
        ✅ Knowledge Base —
        23 Dubai Code Chapters<br>
        ✅ Unlimited document downloads<br>
        ✅ Regular updates & new features
    </div>
</div>
""", unsafe_allow_html=True)

    with col2:
        st.markdown("### 📋 How to Subscribe")
        st.info(
            "**3 Simple Steps:**\n\n"
            "**1️⃣** Choose a payment method below\n\n"
            "**2️⃣** Send **50 AED** (or PKR "
            "equivalent) with your:\n"
            "— Full Name\n"
            "— Email Address\n\n"
            "**3️⃣** Send payment screenshot "
            "via WhatsApp to:\n\n"
            f"📱 **{YOUR_WHATSAPP}**\n\n"
            "✅ Access code sent within **1 hour**\n\n"
            "📅 **30 days full access** activated "
            "immediately on verification"
        )

        st.markdown(f"""
<div style="
    background:#E8F5E9;
    border:2px solid #2E7D32;
    border-radius:10px;
    padding:14px;
    text-align:center;
    margin-top:8px;
">
    <div style="font-weight:700;
                color:#2E7D32;
                font-size:14px;">
        📞 CONTACT FOR SUPPORT
    </div>
    <div style="font-size:14px;
                margin-top:6px;
                color:#1F3864;">
        📱 WhatsApp: <strong>{YOUR_WHATSAPP}</strong>
    </div>
    <div style="font-size:13px;
                color:#555;
                margin-top:4px;">
        ✉️ {YOUR_EMAIL}
    </div>
    <div style="font-size:11px;
                color:#888;
                margin-top:6px;">
        Response within 1 hour
        (Abu Dhabi business hours)
    </div>
</div>
""", unsafe_allow_html=True)

    st.divider()

    # ── PAYMENT METHODS ───────────────────────────────
    st.markdown("## 💳 Payment Methods")

    tab1, tab2, tab3 = st.tabs([
        "🏦 ADCB Bank (UAE — AED)",
        "🏦 Bank Alfalah (Pakistan)",
        "📱 JazzCash (Pakistan)",
    ])

    # ADCB
    with tab1:
        st.markdown(f"""
<div class="pay-method">
    <h3>🏦 ADCB Bank — UAE (Pay in AED)</h3>
    <div class="pay-row">
        <span class="pay-label">Bank Name</span>
        <span class="pay-value">
            {ADCB_BANK_NAME}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Account Title</span>
        <span class="pay-value">
            {ADCB_ACCOUNT_TITLE}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Account Number</span>
        <span class="pay-value">
            {ADCB_ACCOUNT_NO}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">IBAN</span>
        <span class="pay-value">
            {ADCB_IBAN}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Amount</span>
        <span class="pay-value"
              style="color:#C00000;
                     font-size:16px;">
            50 AED
        </span>
    </div>
    <div class="important-note">
        ✅ After transfer, send payment
        screenshot via WhatsApp to:
        <strong>{YOUR_WHATSAPP}</strong><br>
        📝 Include your
        <strong>Full Name + Email Address</strong>
        in the message.<br>
        ✉️ Or email receipt to:
        <strong>{YOUR_EMAIL}</strong>
    </div>
</div>
""", unsafe_allow_html=True)

    # Alfalah
    with tab2:
        st.markdown(f"""
<div class="pay-method">
    <h3>🏦 Bank Alfalah — Pakistan
    (Pay in PKR)</h3>
    <div class="pay-row">
        <span class="pay-label">Bank Name</span>
        <span class="pay-value">
            {ALFALAH_BANK_NAME}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Account Title</span>
        <span class="pay-value">
            {ALFALAH_TITLE}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Account Number</span>
        <span class="pay-value">
            {ALFALAH_ACCOUNT_NO}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">IBAN</span>
        <span class="pay-value">
            {ALFALAH_IBAN}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Swift Code</span>
        <span class="pay-value">
            {ALFALAH_SWIFT}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Branch</span>
        <span class="pay-value">
            {ALFALAH_BRANCH}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Branch Code</span>
        <span class="pay-value">
            {ALFALAH_BRANCH_CODE}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Amount</span>
        <span class="pay-value"
              style="color:#C00000;
                     font-size:15px;">
            PKR equivalent of 50 AED
            (~3,800–4,200 PKR)
        </span>
    </div>
    <div class="important-note">
        ✅ After transfer, send payment
        screenshot via WhatsApp to:
        <strong>{YOUR_WHATSAPP}</strong><br>
        📝 Include your
        <strong>Full Name + Email Address</strong>
        in the message.<br>
        ✉️ Or email receipt to:
        <strong>{YOUR_EMAIL}</strong>
    </div>
</div>
""", unsafe_allow_html=True)

    # JazzCash
    with tab3:
        st.markdown(f"""
<div class="pay-method">
    <h3>📱 JazzCash — Pakistan (Pay in PKR)</h3>
    <div class="pay-row">
        <span class="pay-label">
            JazzCash Number
        </span>
        <span class="pay-value">
            {JAZZCASH_NUMBER}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Account Name</span>
        <span class="pay-value">
            {JAZZCASH_NAME}
        </span>
    </div>
    <div class="pay-row">
        <span class="pay-label">Amount</span>
        <span class="pay-value"
              style="color:#C00000;
                     font-size:15px;">
            PKR equivalent of 50 AED
            (~3,800–4,200 PKR)
        </span>
    </div>
    <div class="important-note">
        ✅ Open JazzCash app →
        Send Money → Mobile Account<br>
        → Enter <strong>{JAZZCASH_NUMBER}</strong>
        → Enter amount → Send<br><br>
        📸 After payment, send screenshot
        via WhatsApp to:
        <strong>{YOUR_WHATSAPP}</strong><br>
        📝 Include your
        <strong>Full Name + Email Address</strong>
        <br>
        ✉️ Or email to:
        <strong>{YOUR_EMAIL}</strong>
    </div>
</div>
""", unsafe_allow_html=True)

    st.divider()

    # ── ALREADY PAID SECTION ──────────────────────────
    st.markdown("### ✅ Already Paid? Enter Your Code")
    st.caption(
        "After verifying your payment, an access "
        "code will be sent to your WhatsApp or email."
    )

    col_a, col_b = st.columns(2)

    with col_a:
        paid_code = st.text_input(
            "Access Code",
            placeholder=(
                "Enter code received via WhatsApp"
            ),
            type="password",
            key="paid_code_input"
        )
        if st.button(
            "🔓 Activate My Subscription",
            type="primary",
            use_container_width=True
        ):
            valid  = get_timed_password()
            master = MASTER_PASSWORD
            if (paid_code == valid or
                    paid_code == master):
                st.session_state.authenticated = True
                st.session_state.auth_type = \
                    "subscription"
                # 30 days
                st.session_state.\
                    subscription_expiry = (
                    time.time() + (30 * 24 * 3600)
                )
                st.success(
                    "🎉 Subscription Activated! "
                    "You now have 30 days full access."
                )
                time.sleep(1.5)
                st.rerun()
            else:
                st.error(
                    "❌ Invalid access code.\n\n"
                    "Please check the code sent to "
                    "your WhatsApp or contact:\n\n"
                    f"📱 {YOUR_WHATSAPP}\n\n"
                    f"✉️ {YOUR_EMAIL}"
                )

    with col_b:
        st.markdown(f"""
<div style="
    background:#E8F0FE;
    border:1px solid #1F3864;
    border-radius:10px;
    padding:16px;
    font-size:13px;
">
    <strong>📋 What happens after payment:</strong>
    <ol style="margin-top:8px;
               padding-left:18px;
               line-height:1.8;">
        <li>Send payment screenshot to
        WhatsApp: <strong>{YOUR_WHATSAPP}</strong>
        </li>
        <li>Include your
        <strong>Name + Email</strong></li>
        <li>Receive your access code
        within <strong>1 hour</strong></li>
        <li>Enter code above to activate
        <strong>30 days access</strong></li>
    </ol>
</div>
""", unsafe_allow_html=True)

    st.divider()

    if st.button(
        "← Back to Login Page",
        key="back_to_login"
    ):
        st.session_state.trial_started = False
        st.session_state.trial_start   = 0
        st.session_state.show_paywall  = False
        st.rerun()


def show_login_page():
    """Main entry/login page with 3 options."""

    st.markdown("""
<style>
.login-hdr {
    background: linear-gradient(
        135deg, #1F3864 0%, #C00000 100%
    );
    border-radius: 16px;
    padding: 40px 32px;
    text-align: center;
    color: white;
    margin-bottom: 30px;
}
.opt-box {
    border-radius: 12px;
    padding: 22px 18px;
    text-align: center;
    min-height: 210px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
}
</style>
""", unsafe_allow_html=True)

    st.markdown(f"""
<div class="login-hdr">
    <div style="font-size:52px;">🦺</div>
    <div style="font-size:26px;
                font-weight:800;
                margin:10px 0 6px 0;
                letter-spacing:0.5px;">
        UAE HSE Compliance App
    </div>
    <div style="font-size:14px; opacity:0.88;">
        Abu Dhabi &amp; Dubai &nbsp;|&nbsp;
        ADOSH-SF v4.0 &nbsp;|&nbsp;
        Dubai Municipality Code
    </div>
    <hr style="border-color:rgba(255,255,255,0.3);
               margin:14px 0;">
    <div style="font-size:13px; opacity:0.8;">
        Developed by:
        <strong>{YOUR_NAME}</strong> |
        {YOUR_TITLE}
    </div>
    <div style="font-size:12px;
                opacity:0.65;
                margin-top:4px;">
        📱 {YOUR_WHATSAPP} &nbsp;|&nbsp;
        ✉️ {YOUR_EMAIL}
    </div>
</div>
""", unsafe_allow_html=True)

    # 3 Options
    c1, c2, c3 = st.columns(3)

    # Free Trial
    with c1:
        st.markdown("""
<div class="opt-box" style="
    background:#E8F5E9;
    border:2px solid #2E7D32;">
    <div style="font-size:44px;">⏱️</div>
    <div style="font-size:18px;
                font-weight:700;
                color:#2E7D32;
                margin:8px 0;">
        15 Min Free Trial
    </div>
    <div style="font-size:12px;
                color:#555;
                line-height:1.6;">
        Try ALL features completely
        free for 15 minutes.<br>
        No payment required.
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
            st.session_state.trial_started = True
            st.session_state.trial_start   = (
                time.time()
            )
            st.rerun()

    # Access Code Login
    with c2:
        st.markdown("""
<div class="opt-box" style="
    background:#E8F0FE;
    border:2px solid #1F3864;">
    <div style="font-size:44px;">🔑</div>
    <div style="font-size:18px;
                font-weight:700;
                color:#1F3864;
                margin:8px 0;">
        Enter Access Code
    </div>
    <div style="font-size:12px;
                color:#555;
                line-height:1.6;">
        Already subscribed?<br>
        Enter your access code
        to login.
    </div>
</div>
""", unsafe_allow_html=True)
        st.markdown(
            "<br>", unsafe_allow_html=True
        )
        code = st.text_input(
            "Code",
            type="password",
            placeholder="Enter your access code...",
            label_visibility="collapsed",
            key="code_box"
        )
        if st.button(
            "🔓 Login",
            use_container_width=True,
            key="login_btn"
        ):
            if (code == get_timed_password() or
                    code == MASTER_PASSWORD):
                st.session_state.authenticated  = True
                st.session_state.auth_type      = (
                    "subscription"
                )
                st.session_state.\
                    subscription_expiry = (
                    time.time() + (30 * 24 * 3600)
                )
                st.success(
                    "✅ Welcome! Loading your app..."
                )
                time.sleep(1)
                st.rerun()
            elif code:
                st.error(
                    "❌ Wrong code. Contact:\n"
                    f"📱 {YOUR_WHATSAPP}"
                )

    # Subscribe
    with c3:
        st.markdown(f"""
<div class="opt-box" style="
    background:#FFF3E0;
    border:2px solid #E65100;">
    <div style="font-size:44px;">💎</div>
    <div style="font-size:18px;
                font-weight:700;
                color:#E65100;
                margin:8px 0;">
        Subscribe Now
    </div>
    <div style="font-size:26px;
                font-weight:800;
                color:#E65100;">
        50 AED
        <span style="font-size:14px;
                     font-weight:400;">
            /month
        </span>
    </div>
    <div style="font-size:12px;
                color:#555;
                line-height:1.6;">
        JazzCash &nbsp;·&nbsp;
        Bank Alfalah &nbsp;·&nbsp;
        ADCB
    </div>
</div>
""", unsafe_allow_html=True)
        st.markdown(
            "<br>", unsafe_allow_html=True
        )
        if st.button(
            "💳 Subscribe — 50 AED/month",
            use_container_width=True,
            key="subscribe_now"
        ):
            st.session_state.show_paywall = True
            st.rerun()

    # Bottom contact bar
    st.divider()
    st.markdown(f"""
<div style="
    background:#1F3864;
    border-radius:10px;
    padding:12px 20px;
    text-align:center;
    color:white;
    font-size:13px;
">
    📞 <strong>WhatsApp:</strong>
    {YOUR_WHATSAPP}
    &nbsp;&nbsp;|&nbsp;&nbsp;
    ✉️ <strong>Email:</strong>
    {YOUR_EMAIL}
    &nbsp;&nbsp;|&nbsp;&nbsp;
    ⏰ Response within 1 hour
    (Abu Dhabi business hours)
</div>
""", unsafe_allow_html=True)