"""
auth.py — a PIN gate.

Deliberately simple: this keeps a passer-by out of the tool, nothing more.
It is not security. Anyone with the repo can read the PIN, and the session
state it relies on lives in the browser session.

The PIN is read from Streamlit secrets if present, otherwise from FALLBACK_PIN
below. Put it in secrets so it is not in the repo:

    [auth]
    pin = "4417"
"""

import hmac

import streamlit as st

FALLBACK_PIN = "1234"      # used only when secrets has no [auth] pin
MAX_TRIES = 5


def _pin():
    try:
        return str(st.secrets["auth"]["pin"])
    except Exception:
        return FALLBACK_PIN


def require_pin(title="Al Madina stock tools"):
    """Call once at the top of the entry script. Stops the run until the
    right PIN is entered."""
    if st.session_state.get("authed"):
        return

    st.title(title)
    st.caption("Enter the PIN to continue.")

    tries = st.session_state.get("pin_tries", 0)
    if tries >= MAX_TRIES:
        st.error("Too many attempts. Reload the page to try again.")
        st.stop()

    with st.form("pin_form", border=True):
        entered = st.text_input("PIN", type="password",
                                label_visibility="collapsed",
                                placeholder="PIN")
        ok = st.form_submit_button("Enter", type="primary",
                                   use_container_width=True)

    if ok:
        # hmac.compare_digest avoids leaking the answer through timing.
        # Honestly irrelevant at this scale, but it costs nothing.
        if entered and hmac.compare_digest(str(entered), _pin()):
            st.session_state["authed"] = True
            st.session_state["pin_tries"] = 0
            st.rerun()
        else:
            st.session_state["pin_tries"] = tries + 1
            left = MAX_TRIES - st.session_state["pin_tries"]
            st.error(f"Wrong PIN. {left} attempt(s) left."
                     if left > 0 else "Wrong PIN.")

    st.stop()


def sign_out_button(target=None):
    """Optional. Drop it anywhere you want a way out."""
    t = target or st
    if t.button("Sign out", use_container_width=True):
        st.session_state["authed"] = False
        st.rerun()
