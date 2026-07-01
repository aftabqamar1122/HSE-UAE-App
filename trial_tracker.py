import hashlib
import json
import os
import time
from datetime import datetime

# ── WHERE WE STORE TRIAL RECORDS ─────────────────────────
# This file persists on the server between sessions
TRIAL_FILE = "/tmp/hse_used_trials.json"


def load_trials():
    """Load all used trial records."""
    if not os.path.exists(TRIAL_FILE):
        return {}
    try:
        with open(TRIAL_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def save_trials(data):
    """Save trial records to file."""
    try:
        with open(TRIAL_FILE, "w") as f:
            json.dump(data, f)
    except Exception:
        pass


def get_device_fingerprint(request_headers=None):
    """
    Create a unique ID for this browser/device.
    Uses combination of available identifiers.
    """
    import streamlit as st

    # Use Streamlit's session ID as base
    # (this changes per browser session)
    # We combine with IP if available
    session_id = ""

    try:
        # Try to get session context
        ctx = st._runtime.get_instance()
        if ctx:
            session_id = str(id(ctx))
    except Exception:
        pass

    # Use query params as additional fingerprint
    try:
        params = st.query_params
        session_id += str(params)
    except Exception:
        pass

    return session_id


def mark_trial_used(device_id):
    """Mark this device as having used the trial."""
    trials = load_trials()
    trials[device_id] = {
        "used_at": datetime.utcnow().isoformat(),
        "timestamp": time.time()
    }
    save_trials(trials)


def has_used_trial(device_id):
    """Check if this device already used the trial."""
    if not device_id:
        return False
    trials = load_trials()
    return device_id in trials