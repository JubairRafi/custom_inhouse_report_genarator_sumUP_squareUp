from __future__ import annotations

import hashlib
import hmac
import pathlib
import time

import streamlit as st

USERS_FILE = pathlib.Path(__file__).parent / "users.txt"
SESSION_MAX_SECONDS = 20 * 60  # hard cap from login time


def _hash(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _load_users() -> dict[str, str]:
    """Parse users.txt into {username: sha256hex}. Missing file -> {}."""
    users: dict[str, str] = {}
    if not USERS_FILE.exists():
        return users
    for raw in USERS_FILE.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        username, pw_hash = line.split(":", 1)
        users[username.strip()] = pw_hash.strip()
    return users


def verify(username: str, password: str) -> bool:
    stored = _load_users().get(username.strip())
    if stored is None:
        return False
    return hmac.compare_digest(stored, _hash(password))


def _clear_session() -> None:
    for key in ("auth_ok", "auth_user", "auth_login_ts"):
        st.session_state.pop(key, None)


def _render_login(expired: bool = False) -> None:
    # Hide the sidebar (and its page navigation) on the login screen.
    st.markdown(
        "<style>[data-testid='stSidebar'], [data-testid='stSidebarCollapsedControl']"
        " { display: none !important; }</style>",
        unsafe_allow_html=True,
    )
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        st.markdown("### 🔒 Sign in")
        if expired:
            st.info("Session expired — please log in again.")
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log in", use_container_width=True)
        if submitted:
            if verify(username, password):
                st.session_state["auth_ok"] = True
                st.session_state["auth_user"] = username.strip()
                st.session_state["auth_login_ts"] = time.time()
                st.rerun()
            else:
                st.error("Invalid username or password.")
    st.stop()


def require_login() -> None:
    """Gate the current page. Renders a login form and stops if not authed."""
    if st.session_state.get("auth_ok"):
        elapsed = time.time() - st.session_state.get("auth_login_ts", 0)
        if elapsed > SESSION_MAX_SECONDS:
            _clear_session()
            _render_login(expired=True)
            return  # unreachable (st.stop), kept for clarity
        # Authenticated and within window — offer logout, then continue.
        with st.sidebar:
            st.caption(f"Signed in as **{st.session_state.get('auth_user', '')}**")
            if st.button("Log out", use_container_width=True):
                _clear_session()
                st.rerun()
        return

    _render_login()
