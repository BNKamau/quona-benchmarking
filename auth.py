"""
Google OAuth 2.0 authentication for Quona Portfolio Benchmarking Platform.

Setup instructions
──────────────────
1. Go to https://console.cloud.google.com
2. Create a new project (or select an existing one)
3. APIs & Services → OAuth consent screen
   - User Type: "Internal" locks access to your Google Workspace org automatically
   - Fill in app name, user support email, developer contact info
4. APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID
   - Application type: Web application
   - Name: Quona Benchmarking
   - Authorised redirect URIs:
       https://your-app.streamlit.app      ← Streamlit Cloud
       http://localhost:8501               ← local dev
5. Copy the Client ID and Client Secret into Streamlit Cloud app secrets
   (or .streamlit/secrets.toml locally — see secrets.toml.example):
       GOOGLE_CLIENT_ID     = "….apps.googleusercontent.com"
       GOOGLE_CLIENT_SECRET = "GOCSPX-…"
       ALLOWED_DOMAIN       = "quona.com"
       REDIRECT_URI         = "https://your-app.streamlit.app"

Session persistence
───────────────────
Authentication is stored in st.session_state, which survives Streamlit re-runs
and page refreshes within the same browser tab (WebSocket session). Opening a new
tab or closing and reopening the browser requires signing in again. This is
appropriate for an internal tool — add cookie-based persistence if longer-lived
sessions are required.
"""

import urllib.parse

import requests
import streamlit as st

# ── Google OAuth endpoints ─────────────────────────────────────────────────────
_AUTH_URL     = "https://accounts.google.com/o/oauth2/v2/auth"
_TOKEN_URL    = "https://oauth2.googleapis.com/token"
_USERINFO_URL = "https://www.googleapis.com/oauth2/v3/userinfo"
_SCOPES       = "openid email profile"

# ── Session state keys ─────────────────────────────────────────────────────────
_USER_KEY = "_quona_auth_user"
_ERR_KEY  = "_quona_auth_error"

# ── Brand palette (mirrors app.py constants so auth.py is self-contained) ──────
_BG     = "#EFF0EA"
_BLACK  = "#2C2C2A"
_WHITE  = "#FFFFFF"
_GREEN  = "#D5FA94"
_BORDER = "#D4D5CE"
_MUTED  = "#888884"
_WARN   = "#E65100"
_WARN_BG = "#FFF3E0"


# ── Config ─────────────────────────────────────────────────────────────────────
def _cfg() -> dict:
    return {
        "client_id":      st.secrets.get("GOOGLE_CLIENT_ID", ""),
        "client_secret":  st.secrets.get("GOOGLE_CLIENT_SECRET", ""),
        "redirect_uri":   st.secrets.get("REDIRECT_URI", "http://localhost:8501"),
        "allowed_domain": st.secrets.get("ALLOWED_DOMAIN", "quona.com"),
    }


# ── Public API ─────────────────────────────────────────────────────────────────

def is_authenticated() -> bool:
    """Returns True when a valid @quona.com user is present in session state."""
    user = st.session_state.get(_USER_KEY)
    if not user:
        return False
    return user.get("email", "").lower().endswith(f"@{_cfg()['allowed_domain']}")


def current_user() -> dict:
    """Returns the authenticated user's info dict (name, email, picture, …)."""
    return st.session_state.get(_USER_KEY, {})


def logout() -> None:
    st.session_state.pop(_USER_KEY, None)
    st.session_state.pop(_ERR_KEY, None)
    st.rerun()


def render_login_page() -> None:
    """
    Display the sign-in page and process any pending OAuth callback.
    Call st.stop() immediately after this when the user is not authenticated.
    """
    _handle_oauth_callback()
    if is_authenticated():
        st.rerun()
        return
    _render_login_ui()


def render_user_sidebar() -> None:
    """Render the logged-in user's avatar, name, email, and a sign-out button
    at the bottom of the sidebar."""
    user    = current_user()
    name    = user.get("name", "Quona User")
    email   = user.get("email", "")
    picture = user.get("picture", "")

    with st.sidebar:
        st.markdown(
            f"<div style='border-top:1px solid {_BORDER};"
            f"padding-top:14px;margin-top:12px'>"
            f"<div style='display:flex;align-items:center;"
            f"gap:10px;margin-bottom:10px'>"
            + (
                f"<img src='{picture}' style='width:30px;height:30px;"
                f"border-radius:50%;border:1.5px solid {_BORDER};flex-shrink:0'>"
                if picture else ""
            )
            + f"<div style='min-width:0'>"
            f"<div style='font-size:13px;font-weight:600;color:{_BLACK};"
            f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{name}</div>"
            f"<div style='font-size:11px;color:{_MUTED};white-space:nowrap;"
            f"overflow:hidden;text-overflow:ellipsis'>{email}</div>"
            f"</div></div></div>",
            unsafe_allow_html=True,
        )
        if st.button("Sign out", key="_quona_signout"):
            logout()


# ── OAuth internals ────────────────────────────────────────────────────────────

def _build_auth_url() -> str:
    cfg = _cfg()
    params = {
        "client_id":     cfg["client_id"],
        "redirect_uri":  cfg["redirect_uri"],
        "response_type": "code",
        "scope":         _SCOPES,
        "access_type":   "online",
        "prompt":        "select_account",
    }
    return _AUTH_URL + "?" + urllib.parse.urlencode(params)


def _exchange_code(code: str) -> dict | None:
    """Exchange an OAuth authorisation code for user profile info."""
    cfg = _cfg()
    try:
        token_r = requests.post(_TOKEN_URL, data={
            "code":          code,
            "client_id":     cfg["client_id"],
            "client_secret": cfg["client_secret"],
            "redirect_uri":  cfg["redirect_uri"],
            "grant_type":    "authorization_code",
        }, timeout=10)
    except requests.RequestException:
        return None
    if not token_r.ok:
        return None

    access_token = token_r.json().get("access_token", "")
    try:
        ui_r = requests.get(
            _USERINFO_URL,
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=10,
        )
    except requests.RequestException:
        return None
    return ui_r.json() if ui_r.ok else None


def _handle_oauth_callback() -> None:
    """Detect a Google OAuth callback (?code=…) in the URL and process it."""
    code = st.query_params.get("code")
    if not code:
        return

    error = st.query_params.get("error")
    st.query_params.clear()          # clean up the URL immediately

    if error:
        st.session_state[_ERR_KEY] = (
            f"Google sign-in was cancelled ({error}). Please try again."
        )
        return

    user_info = _exchange_code(code)
    if not user_info:
        st.session_state[_ERR_KEY] = (
            "Could not retrieve account details from Google. "
            "Please try again or contact your administrator."
        )
        return

    email   = user_info.get("email", "")
    allowed = _cfg()["allowed_domain"]
    if not email.lower().endswith(f"@{allowed}"):
        st.session_state[_ERR_KEY] = (
            f"Access restricted to Quona Capital team members (@{allowed}). "
            f"You signed in as {email}. Please use your Quona email account."
        )
        return

    st.session_state[_USER_KEY] = user_info
    st.session_state.pop(_ERR_KEY, None)


def _render_login_ui() -> None:
    """Render the centred sign-in card."""
    st.markdown("""
    <style>
      #MainMenu, footer, header { visibility: hidden; }
      .block-container { padding-top: 0 !important; max-width: 100% !important; }
    </style>
    """, unsafe_allow_html=True)

    _, mid, _ = st.columns([1, 1.1, 1])
    with mid:
        st.markdown("<div style='height:80px'></div>", unsafe_allow_html=True)

        # Error banner (domain mismatch, cancelled sign-in, etc.)
        err = st.session_state.get(_ERR_KEY)
        if err:
            st.markdown(
                f"<div style='background:{_WARN_BG};border:1px solid {_WARN};"
                f"border-radius:8px;padding:12px 16px;font-size:13px;"
                f"color:{_WARN};margin-bottom:20px;line-height:1.5'>"
                f"⚠&nbsp; {err}</div>",
                unsafe_allow_html=True,
            )

        auth_url = _build_auth_url()

        st.markdown(
            # Card wrapper
            f"<div style='background:{_WHITE};border:1px solid {_BORDER};"
            f"border-radius:16px;padding:56px 48px 48px;"
            f"box-shadow:0 4px 32px rgba(44,44,42,0.09);text-align:center'>"

            # Wordmark
            f"<div style='font-size:28px;font-weight:800;color:{_BLACK};"
            f"letter-spacing:-0.5px;margin-bottom:4px'>Quona Capital</div>"
            f"<div style='font-size:11px;color:{_MUTED};letter-spacing:.5px;"
            f"text-transform:uppercase;margin-bottom:44px'>"
            f"Portfolio Benchmarking Platform</div>"

            f"<hr style='border:none;border-top:1px solid {_BORDER};margin:0 0 36px'>"

            f"<div style='font-size:14px;color:{_MUTED};line-height:1.6;margin-bottom:32px'>"
            f"Sign in with your Quona Google account to access<br>"
            f"portfolio performance data and exit intelligence.</div>"

            # Google sign-in button
            f"<a href='{auth_url}' target='_self' style='text-decoration:none'>"
            f"<div style='display:inline-flex;align-items:center;gap:12px;"
            f"background:{_WHITE};border:1.5px solid {_BORDER};border-radius:8px;"
            f"padding:13px 30px;font-size:14px;font-weight:600;color:{_BLACK};"
            f"box-shadow:0 1px 4px rgba(44,44,42,0.08);cursor:pointer'>"
            # Google G logo SVG
            f"<svg width='18' height='18' viewBox='0 0 18 18' xmlns='http://www.w3.org/2000/svg'>"
            f"<path d='M17.64 9.2c0-.637-.057-1.251-.164-1.84H9v3.481h4.844"
            f"a4.14 4.14 0 0 1-1.796 2.716v2.259h2.908"
            f"c1.702-1.567 2.684-3.875 2.684-6.615z' fill='#4285F4'/>"
            f"<path d='M9 18c2.43 0 4.467-.806 5.956-2.18l-2.908-2.259"
            f"c-.806.54-1.837.86-3.048.86-2.344 0-4.328-1.584-5.036-3.711"
            f"H.957v2.332A8.997 8.997 0 0 0 9 18z' fill='#34A853'/>"
            f"<path d='M3.964 10.71A5.41 5.41 0 0 1 3.682 9c0-.593.102-1.17"
            f".282-1.71V4.958H.957A8.996 8.996 0 0 0 0 9"
            f"c0 1.452.348 2.827.957 4.042l3.007-2.332z' fill='#FBBC05'/>"
            f"<path d='M9 3.58c1.321 0 2.508.454 3.44 1.345l2.582-2.58"
            f"C13.463.891 11.426 0 9 0A8.997 8.997 0 0 0 .957 4.958"
            f"L3.964 7.29C4.672 5.163 6.656 3.58 9 3.58z' fill='#EA4335'/>"
            f"</svg>"
            f"Sign in with Google"
            f"</div></a>"

            f"<div style='font-size:11px;color:{_MUTED};margin-top:28px'>"
            f"Access restricted to @quona.com accounts</div>"

            f"</div>",
            unsafe_allow_html=True,
        )
