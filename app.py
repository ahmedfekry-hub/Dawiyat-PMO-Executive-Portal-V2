
import base64
import hashlib
import hmac
import time
import io
import json
import re
import shutil
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List, Mapping, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


BASE_DIR = Path(__file__).parent
DASHBOARD_PATH = BASE_DIR / "dashboard" / "dashboard.html"
DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = BASE_DIR / "backups"
BACKUP_DIR.mkdir(exist_ok=True)

WO_PATH = DATA_DIR / "u_osp_work_order.csv"
PENALTIES_PATH = DATA_DIR / "Penalties.csv"
DOC_STATUS_CACHE_PATH = DATA_DIR / "Dawiyat_Document_Status.csv"
DISTRICT_PATH = DATA_DIR / "District.csv"

DATA_FILES = {
    "u_osp_work_order.csv": WO_PATH,
    "Penalties.csv": PENALTIES_PATH,
    "District.csv": DISTRICT_PATH,
}

ASSETS_DIR = BASE_DIR / "assets"
DAWIYAT_LOGO_PATH = ASSETS_DIR / "dawiyat_logo.jpg"
MET_LOGO_PATH = ASSETS_DIR / "met_logo.jpg"
PPT_COVER_PATH = ASSETS_DIR / "ppt_cover.png"

DOCUMENT_TYPES = ["Design", "Permit", "Photos", "PAT", "AsBuilt", "Handover", "Commercial"]
DOCUMENT_FOLDER_MAP = {
    "Design": "01 Design",
    "Permit": "02 Permit",
    "Photos": "03 Photos",
    "PAT": "04 PAT",
    "AsBuilt": "05 AsBuilt",
    "Handover": "06 Handover",
    "Commercial": "07 Commercial",
}
DOCUMENT_EXTENSIONS = {
    "Design": ["pdf", "zip", "dwg", "dxf", "xlsx", "xls", "docx"],
    "Permit": ["pdf", "jpg", "jpeg", "png", "xlsx", "xls", "docx"],
    "Photos": ["jpg", "jpeg", "png", "webp", "zip", "rar", "7z"],
    "PAT": ["pdf", "xlsx", "xls", "docx", "jpg", "jpeg", "png"],
    "AsBuilt": ["pdf", "zip", "dwg", "dxf", "xlsx", "xls", "docx"],
    "Handover": ["pdf", "xlsx", "xls", "docx", "jpg", "jpeg", "png", "zip"],
    "Commercial": ["pdf", "xlsx", "xls", "docx", "zip"],
}
GOOGLE_DRIVE_FOLDER_MIME = "application/vnd.google-apps.folder"


def image_to_base64(path: Path) -> str:
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode("ascii")



ROLE_PERMISSIONS = {
    # Full system owner.
    "admin": {
        "dashboard": True, "assistant": True, "alerts": True, "reports": True,
        "admin": True, "upload": True, "email": True, "documents": True,
        "export": True, "export_excel": True, "export_pdf": True,
        "dashboard_tabs": ["overview", "performance", "tables", "decision", "pmo", "perf-explanation"],
    },
    # PMO can manage operational pages and upload CSV/documents, but cannot access Admin Board.
    "pmo": {
        "dashboard": True, "assistant": True, "alerts": True, "reports": True,
        "admin": False, "upload": True, "email": False, "documents": True,
        "export": True, "export_excel": True, "export_pdf": True,
        "dashboard_tabs": ["overview", "performance", "tables", "decision", "pmo", "perf-explanation"],
    },
    # Board sees executive-level content only and PDF export.
    "board": {
        "dashboard": True, "assistant": False, "alerts": False, "reports": True,
        "admin": False, "upload": False, "email": False, "documents": False,
        "export": True, "export_excel": False, "export_pdf": True,
        "dashboard_tabs": ["overview", "decision"],
    },
    # Finance sees SOR/Billing/reporting areas and export, but no upload/admin/audit pages.
    "finance": {
        "dashboard": True, "assistant": False, "alerts": True, "reports": True,
        "admin": False, "upload": False, "email": False, "documents": False,
        "export": True, "export_excel": True, "export_pdf": True,
        "dashboard_tabs": ["tables", "decision"],
    },
    # Viewer is read-only and sees Executive Overview only.
    "viewer": {
        "dashboard": True, "assistant": False, "alerts": False, "reports": False,
        "admin": False, "upload": False, "email": False, "documents": False,
        "export": False, "export_excel": False, "export_pdf": False,
        "dashboard_tabs": ["overview"],
    },
}

ROLE_DISPLAY_NAMES = {
    "admin": "Admin",
    "pmo": "PMO",
    "board": "Board",
    "finance": "Finance",
    "viewer": "Viewer",
}


st.set_page_config(
    page_title="Dawiyat PMO Executive Portal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


PORTAL_CSS = """
<style>
.block-container {
    padding-top: 0.6rem;
    padding-left: 0.7rem;
    padding-right: 0.7rem;
    max-width: 100%;
}
header, footer {visibility: hidden;}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#0f2746 0%,#10223a 100%);
}
[data-testid="stSidebar"] * { color: #f8fafc; }
.portal-login {
    max-width: 1360px;
    width: min(98vw, 1360px);
    margin: 1vh auto 0.8vh;
    padding: 0;
    border-radius: 38px;
    background:
        radial-gradient(circle at top left, rgba(20,184,166,.18), transparent 32%),
        radial-gradient(circle at top right, rgba(245,158,11,.18), transparent 34%),
        linear-gradient(135deg,#ffffff 0%,#f8fbff 100%);
    border: 1px solid #d9e3ef;
    box-shadow: 0 32px 85px rgba(15,23,42,.18);
    overflow: hidden;
    text-align: center;
}
.portal-login-top {
    height: 8px;
    background: linear-gradient(90deg,#f97316,#14b8a6,#2563eb);
}
.portal-login-inner {
    padding: 28px 68px 24px;
}
.logo-row {
    display:grid;
    grid-template-columns: 1.15fr .85fr;
    gap:34px;
    align-items:center;
    margin-bottom:18px;
}
.logo-card {
    height:138px;
    border:1px solid #e6edf5;
    border-radius:28px;
    display:flex;
    align-items:center;
    justify-content:center;
    background:#fff;
    box-shadow: 0 12px 28px rgba(15,23,42,.08);
    padding:18px 26px;
}
.logo-card img {
    max-width:100%;
    max-height:116px;
    object-fit:contain;
}
.logo-card.met-logo img {
    max-height:112px;
}
.portal-badge {
    display:inline-flex;
    padding:8px 16px;
    border-radius:999px;
    background:#10223a;
    color:#facc15;
    font-size:12px;
    font-weight:900;
    letter-spacing:.12em;
    text-transform:uppercase;
    margin-bottom:10px;
}
.portal-title {
    font-size: 48px;
    font-weight: 1000;
    color: #10223a;
    margin-bottom: 8px;
    letter-spacing:-.04em;
}
.portal-subtitle {
    color: #526985;
    font-size: 20px;
    line-height: 1.8;
    max-width:720px;
    margin:0 auto;
}
.portal-feature-row {
    display:grid;
    grid-template-columns: repeat(3,1fr);
    gap:12px;
    margin-top:18px;
}
.portal-feature {
    border:1px solid #e6edf5;
    border-radius:18px;
    background:rgba(255,255,255,.75);
    padding:15px 16px;
    color:#10223a;
    font-size:15px;
    font-weight:800;
}

.account-access-row {
    width: 100%;
    display: flex;
    justify-content: center;
    align-items: center;
    margin: 18px auto 18px;
    text-align: center;
}
.account-access-pill {
    display:flex;
    flex-direction:column;
    align-items:center;
    justify-content:center;
    width:min(1100px,95vw);
    margin: 0 auto;
    padding:18px 40px;
    border-radius:999px;
    background:linear-gradient(135deg,#111827,#1e3a8a);
    color:#facc15 !important;
    font-size:20px;
    line-height:1.45;
    font-weight:1000;
    text-align:center;
    box-shadow:0 14px 28px rgba(15,23,42,.24);
    border:1px solid rgba(255,255,255,.18);
}
.account-access-pill .admin-contact-name {
    display:block;
    color:#ffffff !important;
    font-size:26px;
    font-weight:1000;
    letter-spacing:.01em;
    line-height:1.35;
}
.login-form-wrap {
    max-width: 880px;
    margin: 0 auto 18px;
    padding: 18px 20px 20px;
    border: 6px solid #0b0f19;
    border-radius: 4px;
    background: rgba(255,255,255,.84);
    box-shadow: 0 16px 36px rgba(15,23,42,.14);
}
.login-form-wrap .login-title {
    color: #10223a;
    font-size: 18px;
    font-weight: 1000;
    text-align: center;
    margin-bottom: 10px;
}
div[data-testid="stForm"] {
    max-width: 880px;
    margin: 0 auto 8px;
    padding: 18px 20px 20px;
    border: 6px solid #0b0f19;
    border-radius: 4px;
    background: rgba(255,255,255,.86);
    box-shadow: 0 16px 36px rgba(15,23,42,.14);
}
/* Enlarged login form controls */
div[data-testid="stTextInput"] label {
    font-size: 18px !important;
    font-weight: 900 !important;
    color: #10223a !important;
}
div[data-testid="stTextInput"] input {
    min-height: 54px !important;
    font-size: 18px !important;
    border-radius: 14px !important;
}
div[data-testid="stButton"] button {
    min-height: 50px !important;
    font-size: 18px !important;
    font-weight: 900 !important;
    border-radius: 14px !important;
}

.upload-center-hero {
    border:1px solid #bfdbfe;
    border-radius:18px;
    background:linear-gradient(135deg,#eaf4ff 0%,#f8fbff 100%);
    padding:16px 18px;
    margin:8px 0 10px;
    box-shadow:0 10px 22px rgba(37,99,235,.08);
}
.upload-center-hero .uc-title {
    color:#0f3b73;
    font-size:24px;
    font-weight:1000;
    margin-bottom:5px;
    letter-spacing:-.02em;
}
.upload-center-hero .uc-subtitle {
    color:#315b89;
    font-size:14px;
    line-height:1.6;
    font-weight:700;
}

@media(max-width:760px){
    .portal-login { margin:2vh auto; }
    .portal-login-inner { padding:24px 18px; }
    .logo-row { grid-template-columns:1fr; }
    .portal-title { font-size:30px; }
    .portal-subtitle { font-size:14px; }
    .portal-feature-row { grid-template-columns:1fr; }
}
.exec-card {
    border:1px solid #dbe5f1;
    border-radius:20px;
    background:#fff;
    padding:16px;
    box-shadow:0 10px 24px rgba(15,23,42,.06);
}
.exec-card h3 {margin:0 0 8px;color:#10223a;font-weight:900;}
.exec-card .small {color:#64748b;font-size:13px;}
.metric-tile {
    border:1px solid #dbe5f1;
    border-radius:18px;
    background:#fff;
    padding:16px;
    min-height:120px;
    box-shadow:0 8px 22px rgba(15,23,42,.05);
}
.metric-tile .label {
    color:#64748b;
    font-size:11px;
    font-weight:900;
    text-transform:uppercase;
    letter-spacing:.08em;
}
.metric-tile .value {
    color:#10223a;
    font-size:26px;
    font-weight:900;
    margin-top:8px;
}
.priority-high {background:#fee2e2;color:#991b1b;border-radius:999px;padding:5px 10px;font-weight:900;}
.priority-med {background:#fef3c7;color:#92400e;border-radius:999px;padding:5px 10px;font-weight:900;}
.priority-low {background:#dcfce7;color:#166534;border-radius:999px;padding:5px 10px;font-weight:900;}
.arabic-box {
    direction: rtl;
    text-align: right;
    border:1px solid #dbe5f1;
    border-radius:18px;
    padding:18px;
    background:#fff;
    line-height:1.8;
}
</style>
"""
st.markdown(PORTAL_CSS, unsafe_allow_html=True)


def _default_users() -> Dict[str, Dict[str, str]]:
    """Safe fallback users used only when no valid users are configured in Streamlit Secrets."""
    return {
        "ahmedfekry": {"password": "20020099", "role": "admin"},
        "pmo_team": {"password": "PMO12345", "role": "pmo"},
        "board": {"password": "Met_12345", "role": "board"},
        "finance": {"password": "Finance12345", "role": "finance"},
        "viewer": {"password": "Viewer12345", "role": "viewer"},
    }


def _secret_get(data: Any, key: str, default: str = "") -> str:
    """Read from dict / Streamlit AttrDict / object safely."""
    try:
        if isinstance(data, Mapping):
            return str(data.get(key, default))
        if hasattr(data, "get"):
            return str(data.get(key, default))
        return str(getattr(data, key, default))
    except Exception:
        return str(default)



def _session_secret() -> str:
    """Return a stable secret used only for browser-refresh login persistence.

    Optional Streamlit Secrets:
    [session]
    secret = "PUT_RANDOM_LONG_TEXT_HERE"

    If not configured, the app uses a deterministic fallback from the project name.
    For production, adding [session].secret is recommended but not mandatory.
    """
    try:
        raw = st.secrets.get("session", {})
        secret = _secret_get(raw, "secret", "").strip()
        if secret:
            return secret
    except Exception:
        pass
    try:
        email = _secret_get(st.secrets.get("gcp_service_account", {}), "client_email", "")
        if email:
            return f"dawiyat-pmo-session::{email}"
    except Exception:
        pass
    return "dawiyat-pmo-portal-local-session-secret"


def _make_session_signature(username: str, role: str, password: str, issued_at: str) -> str:
    payload = f"{username}|{role}|{password}|{issued_at}".encode("utf-8")
    return hmac.new(_session_secret().encode("utf-8"), payload, hashlib.sha256).hexdigest()


def _format_login_time(timestamp_str: str | None = None) -> str:
    """Return a readable local login time for the user session."""
    try:
        ts = int(timestamp_str) if timestamp_str else int(time.time())
        return datetime.fromtimestamp(ts).strftime("%d-%b-%Y %I:%M %p")
    except Exception:
        return datetime.now().strftime("%d-%b-%Y %I:%M %p")



def _clear_login_query_params() -> None:
    """Remove persisted login parameters from the URL without touching other app state."""
    try:
        for k in ["auth_user", "auth_role", "auth_iat", "auth_sig"]:
            if k in st.query_params:
                del st.query_params[k]
    except Exception:
        pass


def _persist_login(username: str, role: str, password: str, issued_at: str | None = None) -> None:
    """Persist authenticated login across browser refreshes.

    This prevents pressing the browser Refresh button from sending the user back
    to the login screen. Logout explicitly removes these parameters.
    """
    issued_at = issued_at or str(int(time.time()))
    signature = _make_session_signature(username, role, password, issued_at)
    try:
        st.query_params["auth_user"] = username
        st.query_params["auth_role"] = role
        st.query_params["auth_iat"] = issued_at
        st.query_params["auth_sig"] = signature
    except Exception:
        pass


def _restore_login_from_query_params() -> bool:
    """Restore login state after a browser refresh, if the signed URL session is valid."""
    if st.session_state.get("authenticated"):
        return True
    try:
        username = str(st.query_params.get("auth_user", "")).strip()
        role = str(st.query_params.get("auth_role", "viewer")).strip().lower()
        issued_at = str(st.query_params.get("auth_iat", "")).strip()
        signature = str(st.query_params.get("auth_sig", "")).strip()
    except Exception:
        return False

    if not username or not issued_at or not signature:
        return False

    # Optional safety window: keep browser-refresh sessions valid for 12 hours.
    try:
        if int(time.time()) - int(issued_at) > 12 * 60 * 60:
            _clear_login_query_params()
            return False
    except Exception:
        _clear_login_query_params()
        return False

    users = get_users()
    if username not in users:
        _clear_login_query_params()
        return False

    expected_role = str(users[username].get("role", role)).lower()
    password = str(users[username].get("password", ""))
    expected = _make_session_signature(username, expected_role, password, issued_at)

    if hmac.compare_digest(signature, expected):
        st.session_state["authenticated"] = True
        st.session_state["username"] = username
        st.session_state["role"] = expected_role if expected_role in ROLE_PERMISSIONS else "viewer"
        st.session_state["last_login"] = _format_login_time(issued_at)
        return True

    _clear_login_query_params()
    return False


def get_users() -> Dict[str, Dict[str, str]]:
    """
    Stable role-based user loader.

    Recommended Streamlit Secrets format:

    [users.ahmedfekry]
    password = "20020099"
    role = "admin"

    [users.tamer_solyman]
    password = "Tamer@12345$"
    role = "finance"

    [users.mohamed_syed]
    password = "Mohamed@12345$"
    role = "finance"

    Notes:
    - Users from Secrets are merged with fallback users.
    - If a Secrets user is invalid or missing password, it is ignored.
    - This prevents a bad Secrets entry from breaking all logins.
    """
    users = _default_users()

    try:
        raw_users = st.secrets.get("users", {})
    except Exception:
        return users

    if not raw_users:
        return users

    try:
        items = raw_users.items() if hasattr(raw_users, "items") else []
        for raw_username, data in items:
            username = str(raw_username).strip()
            if not username:
                continue

            if isinstance(data, str):
                password = data.strip()
                role = "viewer"
            else:
                password = _secret_get(data, "password", "").strip()
                role = _secret_get(data, "role", "viewer").strip().lower()

            if not password:
                continue
            if role not in ROLE_PERMISSIONS:
                role = "viewer"

            users[username] = {"password": password, "role": role}
    except Exception:
        # Never block login completely because of a malformed Secrets users section.
        return users

    return users

def current_role() -> str:
    role = str(st.session_state.get("role", "viewer")).lower()
    return role if role in ROLE_PERMISSIONS else "viewer"


def role_policy(role: str | None = None) -> Dict:
    role = (role or current_role()).lower()
    return ROLE_PERMISSIONS.get(role, ROLE_PERMISSIONS["viewer"])


def can(permission: str) -> bool:
    return bool(role_policy().get(permission, False))


def allowed_dashboard_tabs(role: str | None = None) -> List[str]:
    tabs = role_policy(role).get("dashboard_tabs", ["overview"])
    return [str(t) for t in tabs]


def login_page() -> bool:
    if st.session_state.get("authenticated") or _restore_login_from_query_params():
        return True

    dawiyat_logo = image_to_base64(DAWIYAT_LOGO_PATH)
    met_logo = image_to_base64(MET_LOGO_PATH)

    st.markdown(
        f"""
        <div class="portal-login">
            <div class="portal-login-top"></div>
            <div class="portal-login-inner">
                <div class="logo-row">
                    <div class="logo-card dawiyat-logo">
                        <img src="data:image/jpeg;base64,{dawiyat_logo}" alt="Dawiyat Logo" />
                    </div>
                    <div class="logo-card met-logo">
                        <img src="data:image/jpeg;base64,{met_logo}" alt="Middle Sea Telecom Logo" />
                    </div>
                </div>
                <div class="portal-badge">Executive PMO Governance Portal</div>
                <div class="portal-title">Dawiyat PMO Executive Portal</div>
                <div class="portal-subtitle">
                    Secure executive portal for PMO Dashboard, Arabic AI Assistant, Smart Alerts,
                    PDF Reports, CSV Governance, and Executive Decision Support.
                </div>
                <div class="portal-feature-row">
                    <div class="portal-feature">🔐 Role-Based Access</div>
                    <div class="portal-feature">📊 Live PMO Dashboard</div>
                    <div class="portal-feature">🤖 Arabic AI Assistant</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns([0.42, 2.15, 0.42])
    with c2:
        st.markdown(
            """
            <div class="account-access-row">
                <div class="account-access-pill">
                    <div>For account access, please contact</div>
                    <span class="admin-contact-name">Eng./Ahmed Fekry (PMO System Administrator)</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            login = st.form_submit_button("Login", use_container_width=True)

        if login:
            users = get_users()
            if username in users and password == users[username]["password"]:
                role = users[username].get("role", "viewer").lower()
                if role not in ROLE_PERMISSIONS:
                    role = "viewer"
                issued_at = str(int(time.time()))
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.session_state["role"] = role
                st.session_state["last_login"] = _format_login_time(issued_at)
                _persist_login(username, role, users[username]["password"], issued_at=issued_at)
                st.rerun()
            else:
                st.error("Invalid username or password.")

    return False


@st.cache_data(show_spinner=False)
def read_csv_cached(path_str: str, mtime: float) -> pd.DataFrame:
    path = Path(path_str)
    if not path.exists() or path.stat().st_size == 0:
        return pd.DataFrame()
    encodings = ["utf-8-sig", "utf-8", "cp1256", "latin1"]
    last_error = None
    for enc in encodings:
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc)
        except Exception as exc:
            last_error = exc
    st.warning(f"Could not read {path.name}: {last_error}")
    return pd.DataFrame()


def safe_read_csv(path: Path) -> pd.DataFrame:
    mtime = path.stat().st_mtime if path.exists() else 0
    return read_csv_cached(str(path), mtime)


def df_to_records(df: pd.DataFrame) -> List[dict]:
    if df.empty:
        return []
    return df.fillna("").astype(str).to_dict(orient="records")



def read_cached_document_status_records() -> List[dict]:
    """Read last Google Drive document scan status for dashboard preview.

    The HTML dashboard cannot call Google Drive API directly. The Streamlit
    Document Upload Center scans Google Drive and writes this cache so the
    dashboard preview reflects real files inside each Link Code subfolder.
    """
    try:
        if "document_status_df" in st.session_state and isinstance(st.session_state["document_status_df"], pd.DataFrame):
            df = st.session_state["document_status_df"].copy()
            if not df.empty:
                return df_to_records(df)
    except Exception:
        pass
    try:
        if DOC_STATUS_CACHE_PATH.exists():
            df = pd.read_csv(DOC_STATUS_CACHE_PATH, dtype=str).fillna("")
            return df_to_records(df)
    except Exception:
        return []
    return []

def build_initial_raw() -> Dict[str, List[dict]]:
    return {
        "workorders": df_to_records(safe_read_csv(WO_PATH)),
        "penalties": df_to_records(safe_read_csv(PENALTIES_PATH)),
        "districts": df_to_records(safe_read_csv(DISTRICT_PATH)),
        "document_status": read_cached_document_status_records(),
    }


def make_safe_js_initial_raw(raw_data: Dict[str, List[dict]]) -> str:
    payload = json.dumps(raw_data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    b64 = base64.b64encode(payload).decode("ascii")
    return (
        'const INITIAL_RAW = JSON.parse('
        'new TextDecoder().decode('
        f'Uint8Array.from(atob("{b64}"), c => c.charCodeAt(0))'
        '));\n\n'
    )


def inject_data_into_dashboard(html: str, raw_data: Dict[str, List[dict]]) -> str:
    replacement = make_safe_js_initial_raw(raw_data)

    patterns = [
        r"const\s+INITIAL_RAW\s*=\s*.*?;\s*(?=const\s+FILTERS\s*=)",
        r"const\s+INITIAL_RAW\s*=\s*.*?;\s*(?=const\s+state\s*=)",
    ]

    updated = html
    replaced = 0
    for pattern in patterns:
        updated, replaced = re.subn(pattern, replacement, updated, count=1, flags=re.S)
        if replaced:
            break

    if not replaced:
        st.error("CSV injection failed: INITIAL_RAW block was not found inside dashboard.html.")
        return html

    role = current_role()
    policy = role_policy(role)
    allowed_tabs = json.dumps(allowed_dashboard_tabs(role))
    hide_excel = "true" if not policy.get("export_excel", False) else "false"
    hide_pdf = "true" if not policy.get("export_pdf", False) else "false"
    hide_all_exports = "true" if not policy.get("export", False) else "false"
    role_label = ROLE_DISPLAY_NAMES.get(role, role.title())

    portal_patch = f"""
<style>
.file-label, #apply-imports {{ display: none !important; }}
.header-actions::after {{
    content: "Data linked directly from Version 2 Executive Portal";
    display: block;
    color: #64748b;
    font-size: 12px;
    text-align: right;
    margin-top: 4px;
}}
body.role-viewer .doc-repo-panel,
body.role-board .doc-repo-panel,
body.role-finance .doc-repo-panel {{ display:none !important; }}
body.hide-excel button,
body.hide-excel .btn {{}}
</style>
<script>
window.DAWIYAT_RBAC = {{
  role: {json.dumps(role)},
  roleLabel: {json.dumps(role_label)},
  allowedTabs: {allowed_tabs},
  hideExcel: {hide_excel},
  hidePdf: {hide_pdf},
  hideAllExports: {hide_all_exports}
}};
(function applyDawiyatRBAC() {{
  function norm(t) {{ return (t || '').replace(/\s+/g,' ').trim().toLowerCase(); }}
  function hideExportButtons() {{
    const cfg = window.DAWIYAT_RBAC || {{}};
    const buttons = Array.from(document.querySelectorAll('button, a, .btn'));
    buttons.forEach(el => {{
      const txt = norm(el.textContent);
      if (cfg.hideAllExports && (txt.includes('export') || txt.includes('csv') || txt.includes('excel') || txt.includes('pdf'))) {{
        el.style.display = 'none';
        return;
      }}
      if (cfg.hideExcel && (txt.includes('export excel') || txt.includes('export csv') || txt === 'export')) {{
        el.style.display = 'none';
      }}
      if (cfg.hidePdf && txt.includes('export pdf')) {{
        el.style.display = 'none';
      }}
    }});
  }}
  function applyTabs() {{
    const cfg = window.DAWIYAT_RBAC || {{ allowedTabs: ['overview'] }};
    const allowed = new Set(cfg.allowedTabs || ['overview']);
    document.body.classList.add('role-' + (cfg.role || 'viewer'));
    document.querySelectorAll('.tab[data-tab]').forEach(btn => {{
      const ok = allowed.has(btn.dataset.tab);
      btn.style.display = ok ? '' : 'none';
    }});
    ['overview','performance','tables','decision','pmo','perf-explanation'].forEach(tab => {{
      const sec = document.getElementById('tab-' + tab);
      if (sec && !allowed.has(tab)) sec.classList.add('hidden');
    }});
    const active = document.querySelector('.tab.active[data-tab]');
    const activeAllowed = active && allowed.has(active.dataset.tab);
    if (!activeAllowed) {{
      const first = Array.from(allowed)[0] || 'overview';
      if (typeof window.setTab === 'function') window.setTab(first);
      else {{
        document.querySelectorAll('.tab[data-tab]').forEach(t => t.classList.toggle('active', t.dataset.tab === first));
        ['overview','performance','tables','decision','pmo','perf-explanation'].forEach(tab => {{
          const sec = document.getElementById('tab-' + tab);
          if (sec) sec.classList.toggle('hidden', tab !== first);
        }});
      }}
    }}
    hideExportButtons();
  }}
  document.addEventListener('DOMContentLoaded', applyTabs);
  setTimeout(applyTabs, 800);
  setTimeout(applyTabs, 1800);
}})();
</script>
"""
    if "</head>" in updated:
        updated = updated.replace("</head>", portal_patch + "\n</head>", 1)

    return updated


def parse_num(value) -> float:
    if value is None:
        return 0.0
    match = re.search(r"-?\d[\d,]*(?:\.\d+)?", str(value))
    if not match:
        return 0.0
    try:
        return float(match.group(0).replace(",", ""))
    except Exception:
        return 0.0


def first_existing_col(df: pd.DataFrame, candidates: List[str]) -> str:
    normalized = {re.sub(r"[^a-z0-9]", "", c.lower()): c for c in df.columns}
    for c in candidates:
        key = re.sub(r"[^a-z0-9]", "", c.lower())
        if key in normalized:
            return normalized[key]
    return ""


def load_workorders() -> pd.DataFrame:
    wo = safe_read_csv(WO_PATH)
    if wo.empty:
        return wo
    link_col = first_existing_col(wo, ["Link Code"])
    wo_col = first_existing_col(wo, ["Work Order"])
    cost_col = first_existing_col(wo, ["WO Cost", "Cost"])
    progress_col = first_existing_col(wo, ["Percentage of Completion"])
    stage_col = first_existing_col(wo, ["Stage", "Invoice Status", "FULL WO STATUS"])
    sor_col = first_existing_col(wo, ["SOR Status"])
    status_col = first_existing_col(wo, ["Work Order Status"])
    updated_col = first_existing_col(wo, ["Updated"])

    wo["_link"] = wo[link_col].astype(str) if link_col else ""
    wo["_wo"] = wo[wo_col].astype(str) if wo_col else ""
    wo["_cost"] = wo[cost_col].apply(parse_num) if cost_col else 0
    wo["_progress"] = wo[progress_col].apply(parse_num) if progress_col else 0
    wo["_stage"] = wo[stage_col].astype(str) if stage_col else ""
    wo["_sor"] = wo[sor_col].astype(str) if sor_col else ""
    wo["_status"] = wo[status_col].astype(str) if status_col else ""
    wo["_updated"] = wo[updated_col].astype(str) if updated_col else ""
    return wo


def portfolio_metrics() -> Dict[str, float]:
    wo = load_workorders()
    if wo.empty:
        return {"links": 0, "wos": 0, "cost": 0, "progress": 0}
    return {
        "links": wo["_link"].replace("", pd.NA).dropna().nunique(),
        "wos": len(wo),
        "cost": wo["_cost"].sum(),
        "progress": wo["_progress"].mean(),
    }


def render_dashboard() -> None:
    if not DASHBOARD_PATH.exists():
        st.error("Dashboard HTML file is missing: dashboard/dashboard.html")
        return

    raw = build_initial_raw()

    with st.sidebar.expander("Data check", expanded=False):
        st.write(f"Work Orders: {len(raw['workorders']):,}")
        st.write(f"Penalties: {len(raw['penalties']):,}")
        st.write(f"District: {len(raw['districts']):,}")

    # Clear guidance: file upload is a native Streamlit page, not inside the HTML iframe.
    if can("documents"):
        st.markdown(
            """
            <div class="upload-center-hero">
                <div class="uc-title">📤 Document Upload Center</div>
                <div class="uc-subtitle">افتح صفحة رفع الملفات لإدارة مستندات Google Drive حسب كل Link Code، ثم ارجع للداشبورد بدون تسجيل دخول جديد.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("📤 Open Document Upload Center", use_container_width=True, type="secondary"):
            # Do not set st.session_state["main_nav"] here because the sidebar radio
            # with the same key has already been created in this run. Setting it here
            # raises StreamlitAPIException. Use a separate flag and let main() switch
            # the navigation before the radio is instantiated on the next rerun.
            st.session_state["force_document_upload_center"] = True
            st.rerun()

    if can("reports"):
        st.markdown(
            """
            <div class="upload-center-hero">
                <div class="uc-title">📊 Executive PPT Builder</div>
                <div class="uc-subtitle">صفحة مستقلة لتجهيز PowerPoint حسب التقارير المختارة بدون تحميل ثقيل داخل الداشبورد.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("📊 Open Executive PPT Builder", use_container_width=True, type="secondary"):
            st.session_state["force_ppt_builder"] = True
            st.rerun()

    dashboard_html = DASHBOARD_PATH.read_text(encoding="utf-8", errors="ignore")
    dashboard_html = inject_data_into_dashboard(dashboard_html, raw)

    components.html(dashboard_html, height=3100, scrolling=True)


def backup_file(path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = BACKUP_DIR / f"{path.stem}_{stamp}{path.suffix}"
    if path.exists():
        shutil.copy(path, dest)
    return dest


def upload_data_page() -> None:
    st.title("📤 Upload CSV Data")
    st.caption("Upload new CSV files directly from the browser. Existing files are backed up before replacement.")

    if not can("upload"):
        st.error("You do not have permission to upload data.")
        return

    for label, path in DATA_FILES.items():
        with st.container(border=True):
            st.subheader(label)
            current_size = path.stat().st_size / 1024 if path.exists() else 0
            st.caption(f"Current file: {path} | Size: {current_size:,.1f} KB")

            uploaded = st.file_uploader(f"Upload replacement for {label}", type=["csv"], key=f"upload_{label}")

            if uploaded is not None:
                if st.button(f"Apply update: {label}", key=f"apply_{label}", use_container_width=True):
                    backup_file(path)
                    path.write_bytes(uploaded.read())
                    st.cache_data.clear()
                    st.success(f"{label} updated successfully. Previous version was saved in backups.")
                    st.rerun()

    st.divider()
    st.subheader("Previous Versions / Backups")
    backups = sorted(BACKUP_DIR.glob("*"), reverse=True)
    if not backups:
        st.info("No backups yet.")
    else:
        rows = []
        for b in backups:
            rows.append({"File": b.name, "Size KB": round(b.stat().st_size / 1024, 1), "Modified": datetime.fromtimestamp(b.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def smart_alerts_dataframe() -> pd.DataFrame:
    wo = load_workorders()
    if wo.empty:
        return pd.DataFrame()

    alerts = []

    for _, r in wo.iterrows():
        link = r.get("_link", "")
        wo_id = r.get("_wo", "")
        cost = float(r.get("_cost", 0) or 0)
        progress = float(r.get("_progress", 0) or 0)
        stage = str(r.get("_stage", ""))
        sor = str(r.get("_sor", ""))
        status = str(r.get("_status", ""))

        sor_missing = ("created" not in sor.lower()) or ("not" in sor.lower())
        if progress >= 80 and sor_missing:
            alerts.append({
                "Priority": "High",
                "Alert Type": "Revenue Leakage / SOR Missing",
                "Link Code": link,
                "Work Order": wo_id,
                "Cost": cost,
                "Progress": progress,
                "Reason": "Completion >= 80% but SOR is not created.",
                "Required Action": "Commercial / Back Office to create SOR and unblock billing.",
            })

        if "civil" in stage.lower() and ("not" in stage.lower() or "start" in stage.lower()):
            alerts.append({
                "Priority": "High" if cost >= 1_000_000 else "Medium",
                "Alert Type": "Civil Not Start",
                "Link Code": link,
                "Work Order": wo_id,
                "Cost": cost,
                "Progress": progress,
                "Reason": "Civil stage has not started.",
                "Required Action": "Project team to confirm permits, materials, and execution crew readiness.",
            })

        if progress >= 100 and status.lower() != "closed":
            alerts.append({
                "Priority": "Medium",
                "Alert Type": "Completed Not Closed",
                "Link Code": link,
                "Work Order": wo_id,
                "Cost": cost,
                "Progress": progress,
                "Reason": "Work order reached 100% but is not closed.",
                "Required Action": "PMO / Back Office to close administrative and documentation cycle.",
            })

    if not alerts:
        return pd.DataFrame()

    df = pd.DataFrame(alerts)
    order = {"High": 0, "Medium": 1, "Low": 2}
    df["_order"] = df["Priority"].map(order).fillna(9)
    return df.sort_values(["_order", "Cost"], ascending=[True, False]).drop(columns=["_order"])


def smart_alerts_page() -> None:
    st.title("🚨 Smart Alerts Dashboard")
    st.caption("Early-warning executive module based on current uploaded CSV data.")

    df = smart_alerts_dataframe()
    if df.empty:
        st.success("No critical alerts under current data.")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Alerts", f"{len(df):,}")
    c2.metric("High Priority", f"{(df['Priority'] == 'High').sum():,}")
    c3.metric("Exposure Cost", f"{df['Cost'].sum():,.0f}")
    c4.metric("Affected Link Codes", f"{df['Link Code'].nunique():,}")

    priority_filter = st.multiselect("Priority", sorted(df["Priority"].unique()), default=sorted(df["Priority"].unique()))
    type_filter = st.multiselect("Alert Type", sorted(df["Alert Type"].unique()), default=sorted(df["Alert Type"].unique()))
    filtered = df[df["Priority"].isin(priority_filter) & df["Alert Type"].isin(type_filter)]

    st.dataframe(filtered, use_container_width=True, hide_index=True)

    csv = filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download Smart Alerts CSV", csv, "smart_alerts.csv", "text/csv", use_container_width=True)

    if can("email"):
        st.divider()
        st.subheader("Email Alerts")
        recipients = st.text_input("Recipients", placeholder="name@company.com, name2@company.com")
        if st.button("Send Critical Alerts Email", use_container_width=True):
            ok, msg = send_alert_email(recipients, filtered)
            if ok:
                st.success(msg)
            else:
                st.error(msg)


def send_alert_email(recipients: str, df: pd.DataFrame) -> Tuple[bool, str]:
    try:
        smtp_host = st.secrets["email"]["smtp_host"]
        smtp_port = int(st.secrets["email"].get("smtp_port", 587))
        smtp_user = st.secrets["email"]["smtp_user"]
        smtp_password = st.secrets["email"]["smtp_password"]
        sender = st.secrets["email"].get("sender", smtp_user)
    except Exception:
        return False, "Email secrets are not configured. Add [email] settings in Streamlit Secrets."

    recips = [r.strip() for r in recipients.split(",") if r.strip()]
    if not recips:
        return False, "Please enter at least one recipient."

    high = df[df["Priority"] == "High"].head(20)
    body = "Dawiyat PMO Smart Alerts\n\n"
    body += f"Total Alerts: {len(df)}\nHigh Priority: {(df['Priority'] == 'High').sum()}\nExposure Cost: {df['Cost'].sum():,.0f} SAR\n\n"
    body += "Top Critical Alerts:\n"
    for _, r in high.iterrows():
        body += f"- {r['Link Code']} | {r['Alert Type']} | Cost {r['Cost']:,.0f} | Action: {r['Required Action']}\n"

    msg = EmailMessage()
    msg["Subject"] = "Dawiyat PMO Smart Alerts - Critical Sites"
    msg["From"] = sender
    msg["To"] = ", ".join(recips)
    msg.set_content(body)

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

    return True, "Email sent successfully."


def ai_assistant_page() -> None:
    st.title("🤖 AI Executive Assistant بالعربي")
    st.caption("Connected to the latest uploaded CSV files.")

    wo = load_workorders()
    if wo.empty:
        st.warning("No work order data found.")
        return

    metrics = portfolio_metrics()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Link Codes", f"{metrics['links']:,}")
    c2.metric("Work Orders", f"{metrics['wos']:,}")
    c3.metric("WO Cost", f"{metrics['cost']:,.0f}")
    c4.metric("Avg Progress", f"{metrics['progress']:.1f}%")

    q = st.selectbox(
        "اختر سؤال الإدارة التنفيذية",
        [
            "ما أعلى المواقع الجاهزة أو المتقدمة ولم يتم إنشاء SOR لها؟",
            "ما أعلى مواقع Civil Not Start من حيث التكلفة؟",
            "ما الأعمال التي وصلت 100% ولم تغلق؟",
            "ما أكبر عوائق الفوترة الحالية؟",
            "ما المواقع التي يجب تصعيدها اليوم؟",
            "ماذا يحدث إذا لم يتم اتخاذ إجراء خلال 30 يوم؟",
        ],
    )

    alerts = smart_alerts_dataframe()
    st.markdown('<div class="arabic-box">', unsafe_allow_html=True)

    if "SOR" in q:
        df = alerts[alerts["Alert Type"].eq("Revenue Leakage / SOR Missing")] if not alerts.empty else pd.DataFrame()
        st.subheader("Revenue Leakage / SOR Not Created")
        st.write("هذه المواقع متقدمة في الإنجاز ولكنها لا تدخل دورة الفوترة بسبب عدم إنشاء SOR.")
        st.dataframe(df.head(20), use_container_width=True, hide_index=True)
    elif "Civil Not Start" in q:
        df = alerts[alerts["Alert Type"].eq("Civil Not Start")] if not alerts.empty else pd.DataFrame()
        st.subheader("Civil Not Start Escalation")
        st.write("هذه المواقع تحتاج ضغط مباشر على إدارة المشروع لبدء التنفيذ وإزالة عوائق التصاريح أو الموارد أو المواد.")
        st.dataframe(df.head(20), use_container_width=True, hide_index=True)
    elif "100%" in q:
        df = alerts[alerts["Alert Type"].eq("Completed Not Closed")] if not alerts.empty else pd.DataFrame()
        st.subheader("Completed But Not Closed")
        st.write("هذه الأعمال وصلت 100% ولكنها تحتاج إغلاق إداري أو مستندي لتحرير دورة الفوترة.")
        st.dataframe(df.head(20), use_container_width=True, hide_index=True)
    elif "عوائق الفوترة" in q:
        st.subheader("Billing Bottlenecks")
        if not alerts.empty:
            summary = alerts.groupby("Alert Type").agg(Count=("Link Code", "count"), Exposure=("Cost", "sum")).reset_index().sort_values("Exposure", ascending=False)
            st.dataframe(summary, use_container_width=True, hide_index=True)
    elif "تصعيدها" in q:
        st.subheader("Executive Escalation Queue")
        st.write("تم ترتيب هذه القائمة حسب أولوية التعرض التجاري، عدم إنشاء SOR، Civil Not Start، والإغلاق.")
        st.dataframe(alerts.head(20), use_container_width=True, hide_index=True)
    else:
        exposure = alerts["Cost"].sum() if not alerts.empty else 0
        st.subheader("30 Days Executive Impact")
        st.write(f"إذا لم يتم اتخاذ إجراء خلال 30 يوم، ستظل قيمة تقريبية قدرها {exposure:,.0f} SAR معرضة للتجميد أو التأخير في الفوترة والإغلاق.")
        st.write("القرار المطلوب: War Room يومي للفوترة والإغلاق و Civil Not Start.")
    st.markdown("</div>", unsafe_allow_html=True)



def get_drive_service():
    """Create a Google Drive API service using Streamlit Secrets.

    Supports either [google_service_account] or [gcp_service_account]
    because Streamlit examples often use the second name.
    """
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except Exception as exc:
        raise RuntimeError("Google Drive libraries are not installed. Install requirements.txt first.") from exc

    info = None
    for secret_key in ["google_service_account", "gcp_service_account"]:
        try:
            candidate = st.secrets.get(secret_key, None)
            if candidate:
                info = dict(candidate)
                break
        except Exception:
            pass
    if not info:
        raise RuntimeError("Google service account secrets are not configured. Add [google_service_account] or [gcp_service_account] in Streamlit Secrets.")

    if "private_key" in info:
        # Streamlit Secrets uses TOML, while Google exports JSON. Users may paste the key
        # as triple-quoted TOML, or as a JSON-style value containing escaped \\n.
        # Normalize both formats before passing it to google-auth.
        pk = str(info["private_key"]).strip()
        pk = pk.strip('"').strip("'")
        pk = pk.replace("\\r\\n", "\n").replace("\\n", "\n").replace("\r\n", "\n")
        pk = pk.replace("-----BEGIN PRIVATE KEY----- ", "-----BEGIN PRIVATE KEY-----\n")
        pk = pk.replace(" -----END PRIVATE KEY-----", "\n-----END PRIVATE KEY-----")
        pk = re.sub(r"-----BEGIN PRIVATE KEY-----\s*", "-----BEGIN PRIVATE KEY-----\n", pk)
        pk = re.sub(r"\s*-----END PRIVATE KEY-----", "\n-----END PRIVATE KEY-----", pk)
        info["private_key"] = pk
    scopes = ["https://www.googleapis.com/auth/drive"]
    try:
        credentials = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    except Exception as exc:
        raise RuntimeError(
            "Google Drive Service Account private_key is not valid. "
            "In Streamlit Secrets use TOML format. Example: "
            "private_key = \"\"\"-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n\"\"\". "
            "Do not paste JSON with braces or commas."
        ) from exc
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def get_drive_root_folder_id() -> str:
    try:
        return str(st.secrets["google_drive"]["root_folder_id"]).strip()
    except Exception:
        return ""


def google_drive_root_config_message() -> str:
    return (
        "google_drive.root_folder_id is not configured in Streamlit Secrets. "
        "Add the Google Drive folder ID for the main 'Link Codes' folder under [google_drive]. "
        "The current version uses MANUAL Google Drive upload mode: users open the Link Code folder "
        "and upload files directly in Google Drive, while the dashboard scans document status."
    )


def drive_folder_url(folder_id: str) -> str:
    return f"https://drive.google.com/drive/folders/{folder_id}" if folder_id else ""


def extract_drive_folder_id(url_or_id: str) -> str:
    text = str(url_or_id or "").strip()
    if not text:
        return ""
    if re.fullmatch(r"[A-Za-z0-9_-]{20,}", text):
        return text
    for pattern in [r"/folders/([A-Za-z0-9_-]+)", r"[?&]id=([A-Za-z0-9_-]+)", r"/d/([A-Za-z0-9_-]+)"]:
        m = re.search(pattern, text)
        if m:
            return m.group(1)
    return ""


def drive_escape(value: str) -> str:
    return str(value or "").replace("'", "\'")


def find_drive_folder(service, parent_id: str, name: str) -> str:
    query = (
        f"'{drive_escape(parent_id)}' in parents and "
        f"mimeType = '{GOOGLE_DRIVE_FOLDER_MIME}' and "
        f"name = '{drive_escape(name)}' and trashed = false"
    )
    result = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id,name)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = result.get("files", [])
    return files[0]["id"] if files else ""


def create_drive_folder(service, parent_id: str, name: str) -> str:
    metadata = {"name": name, "mimeType": GOOGLE_DRIVE_FOLDER_MIME, "parents": [parent_id]}
    folder = service.files().create(body=metadata, fields="id", supportsAllDrives=True).execute()
    return folder.get("id", "")


def get_or_create_drive_folder(service, parent_id: str, name: str) -> str:
    existing = find_drive_folder(service, parent_id, name)
    return existing or create_drive_folder(service, parent_id, name)


def ensure_link_folder(service, link_code: str, existing_link: str = "") -> tuple:
    """Resolve the existing Link Code folder without uploading/creating files.

    Manual Upload Mode avoids the Service Account storage quota issue.
    It first uses Document_Link from the CSV. If missing, it searches the configured
    Link Codes root folder by Link Code name and stores the found URL back to CSV.
    It does not create folders or upload files.
    """
    folder_id = extract_drive_folder_id(existing_link)
    if folder_id:
        return folder_id, drive_folder_url(folder_id)
    root_id = get_drive_root_folder_id()
    if not root_id:
        raise RuntimeError(google_drive_root_config_message())
    folder_id = find_drive_folder(service, root_id, link_code)
    if folder_id:
        return folder_id, drive_folder_url(folder_id)
    return "", ""


def upload_file_to_drive(service, folder_id: str, uploaded_file, filename: str, mime_type: str = "application/octet-stream") -> str:
    try:
        from googleapiclient.http import MediaIoBaseUpload
    except Exception as exc:
        raise RuntimeError("google-api-python-client is not installed.") from exc
    uploaded_file.seek(0)
    media = MediaIoBaseUpload(uploaded_file, mimetype=mime_type, resumable=True)
    metadata = {"name": filename, "parents": [folder_id]}
    created = service.files().create(
        body=metadata,
        media_body=media,
        fields="id,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return created.get("webViewLink", "")



def safe_drive_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.\-()\[\] ]+", "_", str(value or "")).strip()
    return cleaned or "uploaded_file"


def document_folder_name(doc_type: str) -> str:
    return DOCUMENT_FOLDER_MAP.get(str(doc_type), str(doc_type))


def ensure_document_subfolders(service, link_folder_id: str) -> Dict[str, str]:
    """Return existing standard Option A subfolders under a Link Code folder.

    In Manual Upload Mode the dashboard does not create folders with the
    Service Account because Service Accounts have no personal Drive storage quota.
    Users can create/upload directly inside Google Drive; this function only detects
    existing folders/files and updates the status preview.
    """
    folder_ids = {}
    for doc_type in DOCUMENT_TYPES:
        folder_ids[doc_type] = find_drive_folder(service, link_folder_id, document_folder_name(doc_type))
    return folder_ids


def list_drive_files(service, folder_id: str) -> List[Dict[str, str]]:
    query = f"'{drive_escape(folder_id)}' in parents and mimeType != '{GOOGLE_DRIVE_FOLDER_MIME}' and trashed = false"
    result = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id,name,webViewLink,modifiedTime,size,mimeType)",
        pageSize=100,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        orderBy="modifiedTime desc",
    ).execute()
    return result.get("files", [])


def document_status_for_link(service, link_folder_id: str) -> Dict[str, Dict[str, object]]:
    """Return status/count/latest link for the standard document subfolders."""
    status = {}
    for doc_type in DOCUMENT_TYPES:
        folder_name = document_folder_name(doc_type)
        folder_id = find_drive_folder(service, link_folder_id, folder_name)
        files = list_drive_files(service, folder_id) if folder_id else []
        status[doc_type] = {
            "folder_id": folder_id,
            "folder_url": drive_folder_url(folder_id),
            "count": len(files),
            "latest_file": files[0].get("name", "") if files else "",
            "latest_url": files[0].get("webViewLink", "") if files else "",
            "latest_modified": files[0].get("modifiedTime", "") if files else "",
            "state": "Uploaded" if files else "Missing",
        }
    return status


def status_badge_text(state: str, count: int = 0) -> str:
    if state == "Uploaded":
        return "🟢 Uploaded"
    if state == "Partial":
        return "🟡 Partial"
    return "🔴 Missing"


def get_document_link_for_link(wo: pd.DataFrame, link_code: str) -> str:
    if wo.empty or "Document_Link" not in wo.columns:
        return ""
    vals = wo.loc[wo["Link Code"].astype(str).eq(str(link_code)), "Document_Link"].astype(str)
    vals = [v.strip() for v in vals if str(v).strip()]
    return vals[0] if vals else ""


def update_document_link_in_csv(link_code: str, folder_url: str) -> None:
    wo = safe_read_csv(WO_PATH)
    if wo.empty:
        raise RuntimeError("u_osp_work_order.csv is empty or missing.")
    if "Document_Link" not in wo.columns:
        wo["Document_Link"] = ""
    if "Link Code" not in wo.columns:
        raise RuntimeError("Link Code column is missing from u_osp_work_order.csv.")
    mask = wo["Link Code"].astype(str).eq(str(link_code))
    if not mask.any():
        raise RuntimeError(f"Link Code not found in CSV: {link_code}")
    backup_file(WO_PATH)
    wo.loc[mask, "Document_Link"] = folder_url
    wo.to_csv(WO_PATH, index=False, encoding="utf-8-sig")
    st.cache_data.clear()




def get_all_link_codes(wo: pd.DataFrame) -> List[str]:
    if wo.empty or "Link Code" not in wo.columns:
        return []
    return sorted([x for x in wo["Link Code"].astype(str).str.strip().unique() if x and x.lower() != "nan"])


def build_document_status_rows(service, wo: pd.DataFrame, link_codes: List[str], max_links: int = 50) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Scan Google Drive document folders and return one status row per Link Code.
    The scan is intentionally limited by max_links to avoid slow Google API calls.
    """
    rows = []
    folder_urls = {}
    for link_code in link_codes[:max_links]:
        current_folder_url = get_document_link_for_link(wo, link_code)
        try:
            link_folder_id, link_folder_url = ensure_link_folder(service, link_code, current_folder_url)
            if link_folder_url and not current_folder_url:
                update_document_link_in_csv(link_code, link_folder_url)
            if not link_folder_id:
                raise RuntimeError("No existing Google Drive folder found for this Link Code. Create the folder manually under the Link Codes root folder or add Document_Link to the CSV.")
            status = document_status_for_link(service, link_folder_id)
            folder_urls[link_code] = link_folder_url
            uploaded_count = sum(1 for d in DOCUMENT_TYPES if status.get(d, {}).get("state") == "Uploaded")
            row = {
                "Link Code": link_code,
                "Folder Link": link_folder_url,
                "Uploaded Types": uploaded_count,
                "Missing Types": len(DOCUMENT_TYPES) - uploaded_count,
                "Overall Status": "Uploaded" if uploaded_count == len(DOCUMENT_TYPES) else ("Partial" if uploaded_count > 0 else "Missing"),
                "Latest File": "",
                "Latest Modified": "",
            }
            latest_candidates = []
            for doc_type in DOCUMENT_TYPES:
                info = status.get(doc_type, {})
                row[doc_type] = status_badge_text(str(info.get("state", "Missing")), int(info.get("count", 0) or 0))
                if info.get("latest_modified"):
                    latest_candidates.append((str(info.get("latest_modified", "")), str(info.get("latest_file", ""))))
            if latest_candidates:
                latest_candidates.sort(reverse=True)
                row["Latest Modified"], row["Latest File"] = latest_candidates[0]
            rows.append(row)
        except Exception as exc:
            row = {
                "Link Code": link_code,
                "Folder Link": current_folder_url,
                "Uploaded Types": 0,
                "Missing Types": len(DOCUMENT_TYPES),
                "Overall Status": "Error",
                "Latest File": "",
                "Latest Modified": "",
                "Error": str(exc),
            }
            for doc_type in DOCUMENT_TYPES:
                row[doc_type] = "❌ Missing"
            rows.append(row)
    return pd.DataFrame(rows), folder_urls


def document_kpi_cards(status_df: pd.DataFrame) -> None:
    if status_df.empty:
        st.info("No document status rows scanned yet.")
        return
    total_links = len(status_df)
    complete_links = int((status_df["Overall Status"] == "Uploaded").sum()) if "Overall Status" in status_df.columns else 0
    partial_links = int((status_df["Overall Status"] == "Partial").sum()) if "Overall Status" in status_df.columns else 0
    missing_links = int((status_df["Overall Status"] == "Missing").sum()) if "Overall Status" in status_df.columns else 0
    total_uploaded_types = int(status_df.get("Uploaded Types", pd.Series(dtype=int)).sum())
    total_missing_types = int(status_df.get("Missing Types", pd.Series(dtype=int)).sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Scanned Link Codes", f"{total_links:,}")
    c2.metric("Fully Uploaded", f"{complete_links:,}")
    c3.metric("Partial", f"{partial_links:,}")
    c4.metric("Missing", f"{missing_links:,}")
    c5.metric("Missing Doc Types", f"{total_missing_types:,}", help=f"Uploaded doc types: {total_uploaded_types:,}")


def upload_widget_for_document_type(service, link_code: str, link_folder_id: str, doc_type: str, doc_status: Dict[str, Dict[str, object]]) -> None:
    """Manual upload card: open the correct Google Drive subfolder and scan status.

    Direct upload from Streamlit is intentionally disabled to avoid the Google Drive
    Service Account storage-quota limitation on normal My Drive accounts.
    """
    folder_label = document_folder_name(doc_type)
    info = doc_status.get(doc_type, {})
    with st.container(border=True):
        left, right = st.columns([1, .55])
        with left:
            st.markdown(f"### {folder_label}")
            st.caption(status_badge_text(str(info.get("state", "Missing")), int(info.get("count", 0) or 0)))
            if info.get("latest_file"):
                st.caption(f"Latest: {info.get('latest_file')}")
                st.caption(f"Modified: {info.get('latest_modified', '')}")
        with right:
            if info.get("folder_url"):
                st.link_button("Open Subfolder", str(info.get("folder_url")), use_container_width=True)
            elif link_folder_id:
                st.info("Subfolder not found. Create it manually inside the Link Code folder.")
            if info.get("latest_url"):
                st.link_button("Open Latest File", str(info.get("latest_url")), use_container_width=True)

        st.info("Upload files manually inside Google Drive, then click Refresh Document Status. Direct Streamlit upload is disabled to avoid Service Account storage quota errors.")


def document_upload_page() -> None:
    top_left, top_right = st.columns([1, .22])
    with top_left:
        st.title("📂 Document Upload Center")
        st.caption("Manual Google Drive upload workflow for every Link Code. Open the Link Code folder, upload files directly into: 01 Design / 02 Permit / 03 Photos / 04 PAT / 05 AsBuilt / 06 Handover / 07 Commercial, then refresh status.")
    with top_right:
        st.write("")
        st.write("")
        if st.button("⬅ Back to Dashboard", use_container_width=True):
            st.session_state["force_dashboard"] = True
            st.rerun()

    if not can("documents"):
        st.error("You do not have permission to access documents.")
        return

    wo = safe_read_csv(WO_PATH)
    if wo.empty:
        st.warning("u_osp_work_order.csv is missing or empty.")
        return
    if "Link Code" not in wo.columns:
        st.error("Link Code column is missing from u_osp_work_order.csv.")
        return

    links = get_all_link_codes(wo)
    if not links:
        st.warning("No Link Codes found.")
        return

    try:
        service = get_drive_service()
        drive_connected = True
    except Exception as exc:
        service = None
        drive_connected = False
        st.error(str(exc))
        st.info("Check Streamlit Secrets, enable Google Drive API, and share the Link Codes Google Drive folder with the service account email as Viewer/Editor so the dashboard can scan files.")

    st.markdown("### Executive Documents Dashboard")
    with st.container(border=True):
        scan_col1, scan_col2, scan_col3 = st.columns([1.2, .6, .6])
        with scan_col1:
            scan_scope = st.multiselect("Scan Link Codes", links, default=links[:min(10, len(links))], help="Scanning all Link Codes may take time because each scan checks Google Drive folders.")
        with scan_col2:
            max_scan = st.number_input("Max Scan", min_value=1, max_value=500, value=min(50, len(links)), step=10)
        with scan_col3:
            st.metric("Drive", "Connected" if drive_connected else "Not Connected")
        if drive_connected and st.button("Refresh Document Status", use_container_width=True, type="secondary"):
            with st.spinner("Scanning Google Drive folders..."):
                status_df, folder_urls = build_document_status_rows(service, wo, scan_scope or links, int(max_scan))
                st.session_state["document_status_df"] = status_df
                st.session_state["document_folder_urls"] = folder_urls
                try:
                    status_df.to_csv(DOC_STATUS_CACHE_PATH, index=False, encoding="utf-8-sig")
                    st.cache_data.clear()
                except Exception:
                    pass

        status_df = st.session_state.get("document_status_df", pd.DataFrame())
        document_kpi_cards(status_df)
        if not status_df.empty:
            show_cols = ["Link Code", "Overall Status", "Uploaded Types", "Missing Types"] + DOCUMENT_TYPES + ["Latest File", "Latest Modified", "Folder Link"]
            available_cols = [c for c in show_cols if c in status_df.columns]
            st.dataframe(status_df[available_cols], use_container_width=True, hide_index=True)
            st.download_button(
                "Export Document Status Excel/CSV",
                status_df.to_csv(index=False).encode("utf-8-sig"),
                "Dawiyat_Document_Status.csv",
                "text/csv",
                use_container_width=True,
            )

    st.divider()
    st.markdown("### Open Link Code Folder & Refresh Document Status")
    st.caption("Select one Link Code, open its Google Drive folder, upload files manually in Google Drive, then refresh the status scan. This avoids Service Account storage quota errors.")

    c1, c2, c3 = st.columns([1.35, .75, .75])
    with c1:
        # Keep selected Link Code stable during upload/page reruns.
        if st.session_state.get("doc_upload_link_code") not in links:
            st.session_state["doc_upload_link_code"] = links[0]
        link_code = st.selectbox("Link Code", links, key="doc_upload_link_code")
    with c2:
        st.metric("Document Types", len(DOCUMENT_TYPES))
    with c3:
        st.metric("Mode", "Manual Drive Upload")

    current_folder_url = get_document_link_for_link(wo, link_code)
    link_folder_id = ""
    link_folder_url = current_folder_url
    doc_status = {d: {"state": "Missing", "count": 0, "latest_file": "", "latest_url": "", "folder_url": ""} for d in DOCUMENT_TYPES}

    if drive_connected:
        try:
            link_folder_id, link_folder_url = ensure_link_folder(service, link_code, current_folder_url)
            if link_folder_url and not current_folder_url:
                update_document_link_in_csv(link_code, link_folder_url)
                current_folder_url = link_folder_url
                wo = safe_read_csv(WO_PATH)
            if link_folder_id:
                doc_status = document_status_for_link(service, link_folder_id)
            else:
                st.warning("No existing Google Drive folder found for this Link Code. Create it manually under the Link Codes root folder, or add Document_Link to the CSV.")
        except Exception as exc:
            # This usually means root_folder_id is missing AND the selected Link Code has no Document_Link.
            st.warning(str(exc))
            if current_folder_url:
                st.info("This Link Code has an existing Document_Link, but the folder ID could not be read. Check the link format.")
            else:
                st.info("Select a Link Code that already has Documents = Open, or create a folder with the exact Link Code name under the Link Codes root folder. The dashboard will find it on the next refresh.")

    with st.container(border=True):
        h1, h2 = st.columns([1, .35])
        with h1:
            st.subheader(f"📁 {link_code}")
            st.caption("Manual upload mode: open this folder in Google Drive and upload files into the standard Option A subfolders.")
        with h2:
            if link_folder_url:
                st.link_button("Open Link Code Folder", link_folder_url, use_container_width=True)
            else:
                st.button("No Folder Link", disabled=True, use_container_width=True)

        metrics = st.columns(len(DOCUMENT_TYPES))
        for i, doc_type in enumerate(DOCUMENT_TYPES):
            info = doc_status.get(doc_type, {})
            metrics[i].metric(doc_type, status_badge_text(str(info.get("state", "Missing")), int(info.get("count", 0) or 0)))

    st.warning("Direct Streamlit upload is disabled in this version. Upload files manually inside Google Drive, then click Refresh Document Status.")

    for row_start in range(0, len(DOCUMENT_TYPES), 2):
        cols = st.columns(2)
        for idx, doc_type in enumerate(DOCUMENT_TYPES[row_start:row_start + 2]):
            with cols[idx]:
                upload_widget_for_document_type(service, link_code, link_folder_id, doc_type, doc_status)




def build_pdf_report() -> bytes:
    metrics = portfolio_metrics()
    alerts = smart_alerts_dataframe()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Dawiyat PMO Executive Report", styles["Title"]))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 12))

    summary_data = [
        ["Metric", "Value"],
        ["Link Codes", f"{metrics['links']:,}"],
        ["Work Orders", f"{metrics['wos']:,}"],
        ["WO Cost", f"{metrics['cost']:,.0f} SAR"],
        ["Average Progress", f"{metrics['progress']:.1f}%"],
    ]

    t = Table(summary_data, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#10223a")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#d9e3ef")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("PADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(t)
    story.append(Spacer(1, 16))

    story.append(Paragraph("Top Smart Alerts", styles["Heading2"]))
    if alerts.empty:
        story.append(Paragraph("No critical alerts.", styles["Normal"]))
    else:
        top = alerts.head(12).copy()
        alert_data = [["Priority", "Alert Type", "Link Code", "Cost", "Progress", "Required Action"]]
        for _, r in top.iterrows():
            alert_data.append([
                str(r["Priority"]),
                str(r["Alert Type"])[:36],
                str(r["Link Code"])[:22],
                f"{float(r['Cost']):,.0f}",
                f"{float(r['Progress']):.1f}%",
                str(r["Required Action"])[:70],
            ])
        at = Table(alert_data, colWidths=[60, 160, 120, 80, 70, 300])
        at.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1e3a5f")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#d9e3ef")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("PADDING", (0,0), (-1,-1), 6),
            ("FONTSIZE", (0,0), (-1,-1), 8),
        ]))
        story.append(at)

    story.append(Spacer(1, 14))
    story.append(Paragraph("Prepared by Eng/Ahmed Fekry - Quality & Performance Director", styles["Normal"]))

    doc.build(story)
    return buffer.getvalue()



# -----------------------------------------------------------------------------
# Executive PPT Builder (Streamlit native page)
# This module is intentionally outside dashboard.html to keep the dashboard fast.
# -----------------------------------------------------------------------------
PPT_REPORT_OPTIONS = {
    "kpi_cards": "Executive KPI Cards",
    "portfolio": "Portfolio Summary & Cost Exposure",
    "sor_summary": "SOR Summary & Revenue Exposure",
    "stage_summary": "Overall Stages Summary & Cost Exposure",
    "full_scope": "Dawiyat Project Full Scope",
    "regional": "Regional Performance Summary",
    "completion": "Sites Completion Analysis",
    "cost": "Sites Cost Analysis",
    "monthly": "Monthly Progress Trend",
    "financial": "Executive Financial Report",
    "readiness": "Executive Readiness Report",
}


def _ppt_imports():
    from pptx import Presentation
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    return Presentation, ChartData, XL_CHART_TYPE, XL_LEGEND_POSITION, PP_ALIGN, MSO_ANCHOR, MSO_SHAPE, Inches, Pt, RGBColor


def _clean_text(v: Any, default: str = "N/A") -> str:
    s = str(v if v is not None else "").strip()
    return s if s else default


def _fmt_money(v: float) -> str:
    try:
        return f"{float(v):,.0f}"
    except Exception:
        return "0"


def _fmt_pct(v: float) -> str:
    try:
        return f"{float(v):.0f}%"
    except Exception:
        return "0%"


def _status_from_progress(v: float) -> str:
    try:
        x = float(v)
    except Exception:
        x = 0.0
    if x >= 100:
        return "Completed"
    if x > 0:
        return "In Progress"
    return "Not Start"


def _parse_date_any(v: Any):
    s = str(v or "").strip()
    if not s:
        return pd.NaT
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def load_ppt_workorders() -> pd.DataFrame:
    """Load the same CSV files as the dashboard, but keep it Streamlit-native.
    Includes fields needed for PPT Builder filters and reports.
    """
    wo = safe_read_csv(WO_PATH).copy()
    if wo.empty:
        return pd.DataFrame()

    dist = safe_read_csv(DISTRICT_PATH).copy()

    link_col = first_existing_col(wo, ["Link Code"])
    wo_col = first_existing_col(wo, ["Work Order"])
    cost_col = first_existing_col(wo, ["WO Cost", "Cost"])
    progress_col = first_existing_col(wo, ["Percentage of Completion"])
    subclass_col = first_existing_col(wo, ["Subclass"])
    region_col = first_existing_col(wo, ["Region"])
    city_col = first_existing_col(wo, ["City"])
    district_col = first_existing_col(wo, ["District", "District "])
    project_col = first_existing_col(wo, ["Project"])
    stage_col = first_existing_col(wo, ["Stage"])
    sor_col = first_existing_col(wo, ["SOR Status"])
    sor_ref_col = first_existing_col(wo, ["SOR Reference Number"])
    updated_col = first_existing_col(wo, ["Updated", "Created"])
    closure_col = first_existing_col(wo, ["Work Order Status"])
    performance_col = first_existing_col(wo, ["Performance Status", "Performance"])

    out = pd.DataFrame({
        "Link Code": wo[link_col].astype(str) if link_col else "",
        "Work Order": wo[wo_col].astype(str) if wo_col else "",
        "Region": wo[region_col].astype(str) if region_col else "",
        "City": wo[city_col].astype(str) if city_col else "",
        "District": wo[district_col].astype(str) if district_col else "",
        "Project": wo[project_col].astype(str) if project_col else "",
        "Stage": wo[stage_col].astype(str) if stage_col else "",
        "SOR Status": wo[sor_col].astype(str) if sor_col else "",
        "SOR Reference Number": wo[sor_ref_col].astype(str) if sor_ref_col else "",
        "Cost": wo[cost_col].apply(parse_num) if cost_col else 0.0,
        "Progress": wo[progress_col].apply(parse_num) if progress_col else 0.0,
        "Subclass": wo[subclass_col].astype(str) if subclass_col else "",
        "Updated": wo[updated_col].astype(str) if updated_col else "",
        "Closure Status": wo[closure_col].astype(str) if closure_col else "",
        "Performance": wo[performance_col].astype(str) if performance_col else "",
        "Year": wo[first_existing_col(wo, ["Year"])].astype(str) if first_existing_col(wo, ["Year"]) else "",
        "Work Order Status": wo[first_existing_col(wo, ["Work Order Status"])].astype(str) if first_existing_col(wo, ["Work Order Status"]) else "",
        "Type": wo[first_existing_col(wo, ["Type"])].astype(str) if first_existing_col(wo, ["Type"]) else "",
        "Class": wo[first_existing_col(wo, ["Class"])].astype(str) if first_existing_col(wo, ["Class"]) else "",
        "Scope Target": wo[first_existing_col(wo, ["Scope Target"])].astype(str) if first_existing_col(wo, ["Scope Target"]) else "",
    })

    if not dist.empty:
        d_link = first_existing_col(dist, ["Link Code"])
        d_wo = first_existing_col(dist, ["Work Order"])
        d_region = first_existing_col(dist, ["Region"])
        d_city = first_existing_col(dist, ["City"])
        d_district = first_existing_col(dist, ["District", "District "])
        cols = []
        for c in [d_link, d_wo, d_region, d_city, d_district]:
            if c and c not in cols:
                cols.append(c)
        if d_link and d_wo and cols:
            d = dist[cols].copy()
            rename = {d_link: "Link Code", d_wo: "Work Order"}
            if d_region: rename[d_region] = "Region_map"
            if d_city: rename[d_city] = "City_map"
            if d_district: rename[d_district] = "District_map"
            d = d.rename(columns=rename)
            out = out.merge(d, on=["Link Code", "Work Order"], how="left")
            if "Region_map" in out:
                out["Region"] = out["Region_map"].where(out["Region_map"].astype(str).str.strip().ne(""), out["Region"])
            if "City_map" in out:
                out["City"] = out["City_map"].where(out["City_map"].astype(str).str.strip().ne(""), out["City"])
            if "District_map" in out:
                out["District"] = out["District_map"].where(out["District_map"].astype(str).str.strip().ne(""), out["District"])
            out = out.drop(columns=[c for c in ["Region_map", "City_map", "District_map"] if c in out.columns])

    for c in ["Region", "City", "District", "Project", "Stage", "SOR Status", "SOR Reference Number", "Year", "Work Order Status", "Type", "Class", "Scope Target", "Subclass"]:
        out[c] = out[c].fillna("").astype(str).str.strip()
        out[c] = out[c].replace({"": "N/A", "nan": "N/A", "NaN": "N/A", "None": "N/A"})
    out["Status"] = out["Progress"].apply(_status_from_progress)
    out["Updated_dt"] = out["Updated"].apply(_parse_date_any)
    return out


def link_level_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return pd.DataFrame(columns=["Link Code", "Region", "City", "Cost", "Progress", "Status"])
    grouped = rows.groupby("Link Code", dropna=False).agg(
        Region=("Region", lambda s: _clean_text(next((x for x in s if str(x).strip()), "N/A"))),
        City=("City", lambda s: _clean_text(next((x for x in s if str(x).strip()), "N/A"))),
        Cost=("Cost", "sum"),
        Progress=("Progress", "mean"),
    ).reset_index()
    grouped["Status"] = grouped["Progress"].apply(_status_from_progress)
    return grouped


def city_summary_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    links = link_level_dataframe(rows)
    if links.empty:
        return pd.DataFrame(columns=["Region", "City", "No. of Link Codes", "Completed", "In Progress", "Not Start", "Completion %", "WO Amount"])
    summary = links.groupby(["Region", "City"], dropna=False).agg(
        **{"No. of Link Codes": ("Link Code", "nunique"), "WO Amount": ("Cost", "sum")}
    ).reset_index()
    for status in ["Completed", "In Progress", "Not Start"]:
        counts = links[links["Status"].eq(status)].groupby(["Region", "City"])["Link Code"].nunique()
        summary[status] = summary.set_index(["Region", "City"]).index.map(counts).fillna(0).astype(int)
    summary["Completion %"] = summary.apply(lambda r: (r["Completed"] / r["No. of Link Codes"] * 100) if r["No. of Link Codes"] else 0, axis=1)
    return summary.sort_values("WO Amount", ascending=False)


def status_cost_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    links = link_level_dataframe(rows)
    if links.empty:
        return pd.DataFrame(columns=["Status", "CIVIL", "WO Cost Civil", "FIBRE", "WO Cost Fibre", "Total WO Cost"])
    result = []
    for st_name in ["Completed", "In Progress", "Not Start"]:
        st_rows = rows[rows["Status"].eq(st_name)]
        civil = st_rows[st_rows["Subclass"].str.lower().str.contains("civil", na=False)]
        fiber = st_rows[st_rows["Subclass"].str.lower().str.contains("fiber", na=False)]
        result.append({
            "Status": st_name,
            "CIVIL": civil["Link Code"].nunique(),
            "WO Cost Civil": civil["Cost"].sum(),
            "FIBRE": fiber["Link Code"].nunique(),
            "WO Cost Fibre": fiber["Cost"].sum(),
            "Total WO Cost": st_rows["Cost"].sum(),
        })
    return pd.DataFrame(result)


def monthly_progress_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty or rows["Updated_dt"].isna().all():
        return pd.DataFrame(columns=["Month", "Civil %", "Fiber %", "Overall %"])
    temp = rows.dropna(subset=["Updated_dt"]).copy()
    temp["Month_dt"] = temp["Updated_dt"].dt.to_period("M").dt.to_timestamp()
    result = []
    for month, grp in temp.groupby("Month_dt"):
        civil = grp[grp["Subclass"].str.lower().str.contains("civil", na=False)]
        fiber = grp[grp["Subclass"].str.lower().str.contains("fiber", na=False)]
        result.append({
            "Month": month.strftime("%B %Y"),
            "Civil %": civil["Progress"].mean() if not civil.empty else 0,
            "Fiber %": fiber["Progress"].mean() if not fiber.empty else 0,
            "Overall %": grp["Progress"].mean() if not grp.empty else 0,
        })
    return pd.DataFrame(result).tail(9)


def penalty_total_filtered(rows: pd.DataFrame) -> float:
    pen = safe_read_csv(PENALTIES_PATH)
    if pen.empty:
        return 0.0
    link_col = first_existing_col(pen, ["Cluster Name", "Link Code"])
    amt_col = first_existing_col(pen, ["Penalties Amount", "Penalty Amount"])
    if not link_col or not amt_col:
        return 0.0
    valid_links = set(rows["Link Code"].dropna().astype(str)) if not rows.empty else set()
    p = pen.copy()
    p["_link"] = p[link_col].astype(str).str.strip()
    p["_amount"] = p[amt_col].apply(parse_num)
    p = p[p["_link"].ne("")]
    p = p[~p["_link"].str.lower().isin(["grand total", "total"])]
    if valid_links:
        p = p[p["_link"].isin(valid_links)]
    return float(p["_amount"].sum())


def readiness_summary_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    links = set(rows["Link Code"].dropna().astype(str)) if not rows.empty else set()
    status_df = pd.DataFrame(read_cached_document_status_records())
    out = []
    for doc in DOCUMENT_TYPES:
        uploaded = partial = missing = 0
        if status_df.empty or "Link Code" not in status_df.columns:
            missing = len(links)
        else:
            df = status_df[status_df["Link Code"].astype(str).isin(links)].copy() if links else status_df.copy()
            col = doc if doc in df.columns else first_existing_col(df, [doc])
            if not col:
                missing = len(links)
            else:
                s = df[col].astype(str).str.lower()
                uploaded = int(s.str.contains("uploaded|stage data available|accepted", regex=True).sum())
                partial = int(s.str.contains("partial|under|requested", regex=True).sum())
                missing = max(0, len(links) - uploaded - partial)
        out.append({"Document Type": doc, "Uploaded": uploaded, "Partial": partial, "Missing": missing})
    return pd.DataFrame(out)


def _slide_bg(slide, rgb=(253, 226, 204)):
    _, _, _, _, _, _, MSO_SHAPE, Inches, _, RGBColor = _ppt_imports()
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb)
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(*rgb)
    shape.line.fill.background()
    shape.z_order if hasattr(shape, "z_order") else None


def _add_logo(slide, path: Path, x: float, y: float, w: float, h: float):
    _, _, _, _, _, _, _, Inches, _, _ = _ppt_imports()
    if path.exists():
        slide.shapes.add_picture(str(path), Inches(x), Inches(y), width=Inches(w), height=Inches(h))


def _add_header(slide, title: str):
    _, _, _, _, PP_ALIGN, _, _, Inches, Pt, RGBColor = _ppt_imports()
    _slide_bg(slide)
    _add_logo(slide, MET_LOGO_PATH, 0.25, 0.12, 1.95, 0.85)
    _add_logo(slide, DAWIYAT_LOGO_PATH, 10.55, 0.12, 2.55, 0.75)
    box = slide.shapes.add_textbox(Inches(2.1), Inches(0.52), Inches(9.1), Inches(0.42))
    p = box.text_frame.paragraphs[0]
    p.text = title
    p.alignment = PP_ALIGN.CENTER
    run = p.runs[0]
    run.font.bold = True; run.font.size = Pt(20); run.font.color.rgb = RGBColor(0, 0, 0)


def _add_footer(slide):
    """Fixed footer on all slides except the final Thanks slide."""
    _, _, _, _, PP_ALIGN, _, MSO_SHAPE, Inches, Pt, RGBColor = _ppt_imports()
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.3), Inches(7.05), Inches(10.75), Inches(0.32))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 0, 128); bar.line.fill.background()
    box = slide.shapes.add_textbox(Inches(1.35), Inches(7.07), Inches(10.65), Inches(0.24))
    p = box.text_frame.paragraphs[0]
    p.text = "Prepared by Eng/Ahmed Fekry — Quality & Performance Director"
    p.alignment = PP_ALIGN.CENTER
    run = p.runs[0]
    run.font.bold = True; run.font.size = Pt(13); run.font.color.rgb = RGBColor(255, 221, 0)


def _add_table(slide, headers: List[str], rows: List[List[Any]], x: float, y: float, w: float, h: float, font_size: int = 11):
    _, _, _, _, PP_ALIGN, MSO_ANCHOR, _, Inches, Pt, RGBColor = _ppt_imports()
    table = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(x), Inches(y), Inches(w), Inches(h)).table
    for i in range(len(headers)):
        table.columns[i].width = Inches(w / len(headers))
    for j, head in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = str(head)
        cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(244, 122, 42)
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.bold = True; r.font.size = Pt(font_size); r.font.color.rgb = RGBColor(0, 0, 0)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = str(val)
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(252, 231, 215)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(font_size); r.font.color.rgb = RGBColor(0, 0, 0)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    return table


def _add_bar_chart(slide, title: str, labels: List[str], values: List[float], x: float, y: float, w: float, h: float):
    _, _, _, _, PP_ALIGN, _, MSO_SHAPE, Inches, Pt, RGBColor = _ppt_imports()
    box = slide.shapes.add_textbox(Inches(x), Inches(y - 0.25), Inches(w), Inches(0.22))
    p = box.text_frame.paragraphs[0]; p.text = title; p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.bold = True; p.runs[0].font.size = Pt(12); p.runs[0].font.color.rgb = RGBColor(85, 85, 85)
    max_val = max([float(v or 0) for v in values] + [1])
    for i, (label, value) in enumerate(list(zip(labels, values))[:10]):
        yy = y + 0.08 + i * (h / 10)
        slide.shapes.add_textbox(Inches(x), Inches(yy), Inches(1.55), Inches(0.15)).text = str(label)[:18]
        bw = max(0.05, (float(value or 0) / max_val) * (w - 2.2))
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 1.75), Inches(yy), Inches(bw), Inches(0.12))
        bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(79, 134, 217); bar.line.fill.background()
        slide.shapes.add_textbox(Inches(x + 1.8 + bw), Inches(yy - 0.02), Inches(1.1), Inches(0.17)).text = _fmt_money(value)


def _add_pie_chart(slide, title: str, labels: List[str], values: List[float], x: float, y: float, w: float, h: float):
    _, ChartData, XL_CHART_TYPE, XL_LEGEND_POSITION, _, _, _, Inches, _, _ = _ppt_imports()
    chart_data = ChartData()
    chart_data.categories = labels[:8] or ["N/A"]
    chart_data.add_series(title, values[:8] if values else [1])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(x), Inches(y), Inches(w), Inches(h), chart_data).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False



def apply_ppt_filters(rows: pd.DataFrame, filters: Mapping[str, Any]) -> pd.DataFrame:
    """Apply PPT Builder filters. This replaces iframe-to-Streamlit filter sharing, which is not reliable in Streamlit components."""
    if rows.empty or not filters:
        return rows
    out = rows.copy()
    for col, val in filters.items():
        if col not in out.columns:
            continue
        if val is None:
            continue
        vals = val if isinstance(val, (list, tuple, set)) else [val]
        vals = [str(v).strip() for v in vals if str(v).strip() and str(v).strip() != "All"]
        if vals:
            out = out[out[col].astype(str).isin(vals)]
    return out


def _opt_values(df: pd.DataFrame, col: str) -> List[str]:
    if df.empty or col not in df.columns:
        return []
    vals = sorted([str(x) for x in df[col].dropna().astype(str).unique() if str(x).strip() and str(x).strip() not in ["nan", "None"]])
    return vals


def portfolio_summary_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return pd.DataFrame(columns=["Project", "Link Codes", "WOs", "WO Cost", "Avg Progress", "Share"])
    total_cost = rows["Cost"].sum() or 1
    g = rows.groupby("Project", dropna=False).agg(
        **{"Link Codes": ("Link Code", "nunique"), "WOs": ("Work Order", "count"), "WO Cost": ("Cost", "sum"), "Avg Progress": ("Progress", "mean")}
    ).reset_index()
    g["Share"] = g["WO Cost"] / total_cost * 100
    return g.sort_values("WO Cost", ascending=False)


def stage_summary_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return pd.DataFrame(columns=["Stage", "Link Codes", "WOs", "WO Cost", "Avg Progress", "Share"])
    total_cost = rows["Cost"].sum() or 1
    g = rows.groupby("Stage", dropna=False).agg(
        **{"Link Codes": ("Link Code", "nunique"), "WOs": ("Work Order", "count"), "WO Cost": ("Cost", "sum"), "Avg Progress": ("Progress", "mean")}
    ).reset_index()
    g["Share"] = g["WO Cost"] / total_cost * 100
    return g.sort_values("WO Cost", ascending=False)


def sor_summary_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return pd.DataFrame(columns=["SOR Status", "Link Codes", "WO Cost", "Share"])
    links = link_level_dataframe(rows)
    total_cost = links["Cost"].sum() or 1
    g = links.groupby("SOR Status", dropna=False).agg(
        **{"Link Codes": ("Link Code", "nunique"), "WO Cost": ("Cost", "sum")}
    ).reset_index()
    g["Share"] = g["WO Cost"] / total_cost * 100
    return g.sort_values("WO Cost", ascending=False)


def kpi_cards_dataframe(rows: pd.DataFrame) -> pd.DataFrame:
    penalties = penalty_total_filtered(rows)
    links = rows["Link Code"].nunique() if not rows.empty else 0
    cost = rows["Cost"].sum() if not rows.empty else 0
    civil = rows[rows["Subclass"].str.lower().str.contains("civil", na=False)]
    fiber = rows[rows["Subclass"].str.lower().str.contains("fiber", na=False)]
    closed = rows[rows["Closure Status"].astype(str).str.lower().str.contains("closed", na=False)]
    pending = rows[(rows["Progress"] >= 100) & (~rows.index.isin(closed.index))]
    data = [
        ["Total Link Codes", f"{links:,}"],
        ["Total WO Cost", _fmt_money(cost)],
        ["Civil Completion Rate", _fmt_pct(civil["Progress"].mean() if not civil.empty else 0)],
        ["Fiber Completion Rate", _fmt_pct(fiber["Progress"].mean() if not fiber.empty else 0)],
        ["Penalties", _fmt_money(penalties)],
        ["Closed Indicators", f"{closed['Link Code'].nunique():,}"],
        ["Pending Closure", f"{pending['Link Code'].nunique():,}"],
    ]
    return pd.DataFrame(data, columns=["Metric", "Value"])


def _add_cover_slide(prs, blank):
    _, _, _, _, PP_ALIGN, _, MSO_SHAPE, Inches, Pt, RGBColor = _ppt_imports()
    slide = prs.slides.add_slide(blank)
    if PPT_COVER_PATH.exists():
        slide.shapes.add_picture(str(PPT_COVER_PATH), Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
    else:
        _slide_bg(slide, rgb=(16, 34, 58))
        _add_logo(slide, DAWIYAT_LOGO_PATH, 0.6, 0.45, 2.5, 0.9)
        _add_logo(slide, MET_LOGO_PATH, 10.4, 0.45, 2.2, 1.0)
        box = slide.shapes.add_textbox(Inches(1.0), Inches(3.0), Inches(11.3), Inches(0.75))
        p = box.text_frame.paragraphs[0]
        p.text = "Dawiyat Executive Presentation"
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.bold = True; p.runs[0].font.size = Pt(32); p.runs[0].font.color.rgb = RGBColor(255,255,255)
    _add_footer(slide)
    return slide


def _add_thanks_slide(prs, blank):
    _, _, _, _, PP_ALIGN, _, MSO_SHAPE, Inches, Pt, RGBColor = _ppt_imports()
    slide = prs.slides.add_slide(blank)
    _slide_bg(slide)
    title = slide.shapes.add_textbox(Inches(3.3), Inches(2.55), Inches(6.8), Inches(0.9))
    p = title.text_frame.paragraphs[0]
    p.text = "Thanks"
    p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.size = Pt(54); p.runs[0].font.color.rgb = RGBColor(0, 32, 96)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.45), Inches(5.35), Inches(10.45), Inches(0.45))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 0, 128); bar.line.fill.background()
    tx = slide.shapes.add_textbox(Inches(1.55), Inches(5.39), Inches(10.25), Inches(0.35))
    p = tx.text_frame.paragraphs[0]
    p.text = "Prepared by Eng/Ahmed Fekry — Quality & Performance Director"
    p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.bold = True; p.runs[0].font.size = Pt(22); p.runs[0].font.color.rgb = RGBColor(255, 221, 0)
    return slide



def build_ppt_report(selected_reports: List[str], rows: pd.DataFrame | None = None) -> bytes:
    Presentation, *_ = _ppt_imports()
    prs = Presentation()
    prs.slide_width = 12192000  # 13.333 in
    prs.slide_height = 6858000  # 7.5 in
    blank = prs.slide_layouts[6]

    rows = rows.copy() if rows is not None else load_ppt_workorders()
    cities = city_summary_dataframe(rows)
    status_cost = status_cost_dataframe(rows)
    months = monthly_progress_dataframe(rows)
    readiness = readiness_summary_dataframe(rows)
    portfolio = portfolio_summary_dataframe(rows)
    stages = stage_summary_dataframe(rows)
    sor = sor_summary_dataframe(rows)

    # Fixed first page
    _add_cover_slide(prs, blank)

    if "kpi_cards" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Executive KPI Cards")
        kpi = kpi_cards_dataframe(rows)
        _add_table(slide, ["Metric", "Value"], kpi.values.tolist(), 1.45, 1.35, 10.4, 4.6, 15)
        _add_footer(slide)

    if "portfolio" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Portfolio Summary & Cost Exposure")
        t = portfolio.head(10).copy()
        body = [[r["Project"], f'{int(r["Link Codes"]):,}', f'{int(r["WOs"]):,}', _fmt_money(r["WO Cost"]), _fmt_pct(r["Avg Progress"]), _fmt_pct(r["Share"])] for _, r in t.iterrows()]
        _add_table(slide, ["Project", "Link Codes", "WOs", "WO Cost", "Avg Progress", "Share"], body, 0.35, 1.25, 12.6, 3.7, 9)
        _add_bar_chart(slide, "Project Distribution by Cost", t["Project"].astype(str).tolist(), t["WO Cost"].tolist(), 0.8, 5.3, 11.5, 1.35)
        _add_footer(slide)

    if "sor_summary" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "SOR Summary & Revenue Exposure")
        t = sor.head(10).copy()
        body = [[r["SOR Status"], f'{int(r["Link Codes"]):,}', _fmt_money(r["WO Cost"]), _fmt_pct(r["Share"])] for _, r in t.iterrows()]
        _add_table(slide, ["SOR Status", "Link Codes", "WO Cost", "Share"], body, 0.7, 1.25, 5.8, 4.8, 10)
        _add_pie_chart(slide, "SOR Status Distribution", t["SOR Status"].astype(str).tolist(), t["WO Cost"].tolist(), 7.0, 1.35, 5.6, 4.7)
        _add_footer(slide)

    if "stage_summary" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Overall Stages Summary & Cost Exposure")
        t = stages.head(10).copy()
        body = [[r["Stage"], f'{int(r["Link Codes"]):,}', f'{int(r["WOs"]):,}', _fmt_money(r["WO Cost"]), _fmt_pct(r["Avg Progress"]), _fmt_pct(r["Share"])] for _, r in t.iterrows()]
        _add_table(slide, ["Stage", "Link Codes", "WOs", "WO Cost", "Avg Progress", "Share"], body, 0.35, 1.15, 12.6, 3.8, 8)
        _add_bar_chart(slide, "Stage Cost Ranking", t["Stage"].astype(str).tolist(), t["WO Cost"].tolist(), 0.8, 5.2, 11.5, 1.45)
        _add_footer(slide)

    if "full_scope" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Dawiyat Project Full Scope 2025  & 2026")
        t = cities[["Region", "City", "No. of Link Codes", "WO Amount"]].head(10).copy()
        body = [[r["Region"], r["City"], f'{int(r["No. of Link Codes"]):,}', _fmt_money(r["WO Amount"])] for _, r in t.iterrows()]
        body.append(["Grand Total", "", f'{int(cities["No. of Link Codes"].sum()):,}', _fmt_money(cities["WO Amount"].sum())])
        _add_table(slide, ["Region", "City", "No. of Link Codes", "WO Amount"], body, 0.3, 1.55, 6.6, 4.2, 10)
        _add_pie_chart(slide, "City Business Volume", cities["City"].tolist(), cities["WO Amount"].tolist(), 7.05, 1.35, 5.7, 4.7)
        _add_footer(slide)

    if "regional" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Regional Performance Summary")
        t = cities.head(12)
        body = [[r["Region"], r["City"], f'{int(r["No. of Link Codes"]):,}', int(r["Completed"]), int(r["In Progress"]), int(r["Not Start"]), _fmt_pct(r["Completion %"])] for _, r in t.iterrows()]
        _add_table(slide, ["Region", "City", "No. of Link Codes", "Completed", "In Progress", "Not Start", "Completion %"], body, 0.35, 1.25, 12.6, 3.6, 9)
        _add_bar_chart(slide, "City Chart Area", t["City"].tolist(), t["WO Amount"].tolist(), 0.8, 5.25, 11.5, 1.45)
        _add_footer(slide)

    if "completion" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Sites Completion Analysis")
        txt = slide.shapes.add_textbox(Inches(0.25), Inches(1.25), Inches(12.8), Inches(0.75))
        txt.text = "The Site Completion Analysis provides a comprehensive overview of project progress by evaluating the status of all planned FTTH sites under the current filtered data scope."
        total_links = max(1, link_level_dataframe(rows)["Link Code"].nunique())
        body = []
        for _, r in status_cost.iterrows():
            body.append([r["Status"], int(r["CIVIL"]), _fmt_pct(int(r["CIVIL"]) / total_links * 100), int(r["FIBRE"]), _fmt_pct(int(r["FIBRE"]) / total_links * 100), _fmt_pct((int(r["CIVIL"]) + int(r["FIBRE"])) / (2 * total_links) * 100)])
        _add_table(slide, ["Status", "CIVIL", "CIVIL%", "FIBRE", "FIBRE%", "OVERALL%"], body, 1.3, 2.45, 10.4, 2.15, 12)
        _add_pie_chart(slide, "Site Completion Analysis", status_cost["Status"].tolist(), status_cost["Total WO Cost"].tolist(), 4.2, 4.85, 4.5, 1.55)
        _add_footer(slide)

    if "cost" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Sites Cost Analysis")
        txt = slide.shapes.add_textbox(Inches(0.25), Inches(1.25), Inches(12.8), Inches(0.75))
        txt.text = "The Site Cost Analysis provides a detailed assessment of project expenditures and cost efficiency across all implemented FTTH sites under the current filtered data scope."
        body = [[r["Status"], int(r["CIVIL"]), _fmt_money(r["WO Cost Civil"]), int(r["FIBRE"]), _fmt_money(r["WO Cost Fibre"]), _fmt_money(r["Total WO Cost"])] for _, r in status_cost.iterrows()]
        _add_table(slide, ["Status", "CIVIL", "WO Cost", "FIBRE", "WO Cost", "Total WO Cost"], body, 1.2, 2.2, 10.8, 2.35, 12)
        _add_bar_chart(slide, "Cost Analysis", status_cost["Status"].tolist(), status_cost["Total WO Cost"].tolist(), 2.3, 5.2, 8.7, 1.25)
        _add_footer(slide)

    if "monthly" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Monthly Progress Trend 2025-2026")
        txt = slide.shapes.add_textbox(Inches(0.25), Inches(1.32), Inches(12.8), Inches(0.65))
        txt.text = "Monthly progress trend demonstrates advancement across Civil and Fiber implementation activities based on the available Updated dates in the current filtered data scope."
        body = [[r["Month"], _fmt_pct(r["Civil %"]), _fmt_pct(r["Fiber %"]), _fmt_pct(r["Overall %"])] for _, r in months.iterrows()]
        _add_table(slide, ["Month", "Civil %", "Fiber %", "Overall %"], body, 0.65, 2.25, 5.6, 4.2, 11)
        _add_bar_chart(slide, "Overall Monthly Progress", months["Month"].tolist(), months["Overall %"].tolist(), 6.8, 2.35, 5.6, 3.7)
        _add_footer(slide)

    if "financial" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Executive Financial Report")
        penalties = penalty_total_filtered(rows)
        completed_cost = rows.loc[rows["Progress"] >= 100, "Cost"].sum()
        risk_cost = rows.loc[rows["Performance"].str.lower().str.contains("risk|off", na=False), "Cost"].sum()
        pending_closure = rows.loc[(rows["Progress"] >= 100) & (~rows["Closure Status"].str.lower().str.contains("closed", na=False)), "Cost"].sum()
        data = [
            ["Total WO Cost", _fmt_money(rows["Cost"].sum())],
            ["Completed Cost Exposure", _fmt_money(completed_cost)],
            ["Penalty Amount", _fmt_money(penalties)],
            ["At Risk / Off Track Cost", _fmt_money(risk_cost)],
            ["Pending Closure Cost", _fmt_money(pending_closure)],
        ]
        _add_table(slide, ["Metric", "Value"], data, 2.0, 1.55, 9.3, 3.6, 12)
        _add_bar_chart(slide, "Financial Exposure", [x[0] for x in data], [parse_num(x[1]) for x in data], 2.3, 5.35, 8.8, 1.3)
        _add_footer(slide)

    if "readiness" in selected_reports:
        slide = prs.slides.add_slide(blank); _add_header(slide, "Executive Readiness Report")
        if readiness.empty:
            _add_table(slide, ["Document Type", "Uploaded", "Partial", "Missing"], [], 1.3, 1.5, 10.5, 3.5, 12)
        else:
            body = [[r["Document Type"], int(r["Uploaded"]), int(r["Partial"]), int(r["Missing"])] for _, r in readiness.iterrows()]
            _add_table(slide, ["Document Type", "Uploaded", "Partial", "Missing"], body, 1.3, 1.4, 10.6, 4.6, 12)
        _add_footer(slide)

    # Fixed last page - no footer
    _add_thanks_slide(prs, blank)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


def executive_ppt_builder_page() -> None:
    st.title("📊 Executive PPT Builder")
    st.caption("Independent PowerPoint generator using dashboard-aligned filters and selected report slides.")

    if st.button("← Back to Dashboard", use_container_width=True, type="secondary"):
        st.session_state["force_dashboard"] = True
        st.rerun()

    all_rows = load_ppt_workorders()
    if all_rows.empty:
        st.warning("No work order data is available.")
        return

    st.info(
        "Due to Streamlit iframe isolation, the native Streamlit page cannot automatically read filter values selected inside dashboard.html. "
        "For reliability, this page provides the same dashboard filters here and applies them dynamically to the PowerPoint output."
    )

    with st.expander("🎛️ Dashboard Filters for PPT", expanded=True):
        f1, f2, f3 = st.columns(3)

        with f1:
            region_sel = st.multiselect("Region", _opt_values(all_rows, "Region"), default=[])
            city_base = apply_ppt_filters(all_rows, {"Region": region_sel})
            city_sel = st.multiselect("City", _opt_values(city_base, "City"), default=[])
            district_base = apply_ppt_filters(city_base, {"City": city_sel})
            district_sel = st.multiselect("District", _opt_values(district_base, "District"), default=[])
            project_sel = st.multiselect("Project", _opt_values(all_rows, "Project"), default=[])

        with f2:
            stage_sel = st.multiselect("Stage", _opt_values(all_rows, "Stage"), default=[])
            subclass_sel = st.multiselect("Subclass", _opt_values(all_rows, "Subclass"), default=[])
            year_sel = st.multiselect("Year", _opt_values(all_rows, "Year"), default=[])
            wo_status_sel = st.multiselect("Work Order Status", _opt_values(all_rows, "Work Order Status"), default=[])

        with f3:
            sor_sel = st.multiselect("SOR Status", _opt_values(all_rows, "SOR Status"), default=[])
            sor_ref_sel = st.multiselect("SOR Reference Number", _opt_values(all_rows, "SOR Reference Number"), default=[])
            type_sel = st.multiselect("Type", _opt_values(all_rows, "Type"), default=[])
            class_sel = st.multiselect("Class", _opt_values(all_rows, "Class"), default=[])

        scope_sel = st.multiselect("Scope Target", _opt_values(all_rows, "Scope Target"), default=[])

        link_base = apply_ppt_filters(all_rows, {
            "Region": region_sel,
            "City": city_sel,
            "District": district_sel,
            "Project": project_sel,
            "Stage": stage_sel,
            "Subclass": subclass_sel,
            "Year": year_sel,
            "Work Order Status": wo_status_sel,
            "SOR Status": sor_sel,
            "SOR Reference Number": sor_ref_sel,
            "Type": type_sel,
            "Class": class_sel,
            "Scope Target": scope_sel,
        })
        link_sel = st.multiselect("Link Code", _opt_values(link_base, "Link Code")[:1000], default=[])

    ppt_filters = {
        "Region": region_sel,
        "City": city_sel,
        "District": district_sel,
        "Project": project_sel,
        "Stage": stage_sel,
        "Subclass": subclass_sel,
        "Year": year_sel,
        "Work Order Status": wo_status_sel,
        "SOR Status": sor_sel,
        "SOR Reference Number": sor_ref_sel,
        "Type": type_sel,
        "Class": class_sel,
        "Scope Target": scope_sel,
        "Link Code": link_sel,
    }
    rows = apply_ppt_filters(all_rows, ppt_filters)

    metrics_links = rows["Link Code"].nunique()
    metrics_cost = rows["Cost"].sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("Filtered Link Codes", f"{metrics_links:,}")
    c2.metric("Filtered Work Orders", f"{len(rows):,}")
    c3.metric("Filtered WO Cost", f"{metrics_cost:,.0f}")

    selected = st.multiselect(
        "Select reports to include in PowerPoint",
        options=list(PPT_REPORT_OPTIONS.keys()),
        default=list(PPT_REPORT_OPTIONS.keys()),
        format_func=lambda k: PPT_REPORT_OPTIONS[k],
    )

    st.info("Fixed slides are always included: first cover page + last Thanks page. Footer is fixed at the bottom of every slide except the last page.")

    preview_cols = st.columns(2)
    with preview_cols[0]:
        st.subheader("Selected Slide Order")
        order = ["Fixed Cover"] + [PPT_REPORT_OPTIONS[k] for k in selected] + ["Fixed Thanks"]
        st.write(pd.DataFrame({"#": range(1, len(order)+1), "Slide": order}))
    with preview_cols[1]:
        st.subheader("Filtered Data Scope")
        city_summary = city_summary_dataframe(rows)
        st.dataframe(city_summary[["Region", "City", "No. of Link Codes", "WO Amount"]].head(10), use_container_width=True, hide_index=True)

    if rows.empty:
        st.warning("No records match the selected filters.")
        return
    if not selected:
        st.warning("Please select at least one report.")
        return

    if st.button("📊 Generate PowerPoint", use_container_width=True, type="primary"):
        with st.spinner("Generating PowerPoint presentation..."):
            ppt_bytes = build_ppt_report(selected, rows)
        st.download_button(
            "📥 Download Selected PowerPoint Presentation",
            data=ppt_bytes,
            file_name=f"Dawiyat_Executive_Presentation_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )


def reports_page() -> None:
    st.title("📄 Executive Reports")
    st.caption("Download PDF Executive Report based on current CSV data.")

    pdf_bytes = build_pdf_report()
    st.download_button(
        "Download PDF Executive Report",
        pdf_bytes,
        "Dawiyat_PMO_Executive_Report.pdf",
        "application/pdf",
        use_container_width=True,
    )

    st.info("The full HTML dashboard PDF can still be exported from the dashboard Export PDF button.")


def admin_page() -> None:
    st.title("⚙️ Executive Admin Board")
    if not can("admin"):
        st.error("Admin permission required.")
        return

    st.subheader("Users & Roles")
    st.write("Current user:", st.session_state.get("username"))
    st.write("Current role:", ROLE_DISPLAY_NAMES.get(current_role(), current_role().title()))

    users_table = []
    for username, data in get_users().items():
        role = str(data.get("role", "viewer")).lower()
        policy = role_policy(role)
        users_table.append({
            "Username": username,
            "Role": ROLE_DISPLAY_NAMES.get(role, role.title()),
            "Dashboard Tabs": ", ".join(policy.get("dashboard_tabs", [])),
            "Upload CSV": "Yes" if policy.get("upload") else "No",
            "Documents": "Yes" if policy.get("documents") else "No",
            "Export Excel": "Yes" if policy.get("export_excel") else "No",
            "Export PDF": "Yes" if policy.get("export_pdf") else "No",
            "Admin Board": "Yes" if policy.get("admin") else "No",
        })
    st.dataframe(pd.DataFrame(users_table), use_container_width=True, hide_index=True)

    st.subheader("Role Matrix")
    role_rows = []
    for role, policy in ROLE_PERMISSIONS.items():
        role_rows.append({
            "Role": ROLE_DISPLAY_NAMES.get(role, role.title()),
            "Dashboard Pages": ", ".join(policy.get("dashboard_tabs", [])),
            "Upload CSV": "Yes" if policy.get("upload") else "No",
            "Export": "Excel/PDF" if policy.get("export_excel") and policy.get("export_pdf") else ("PDF only" if policy.get("export_pdf") else "No"),
            "Smart Alerts": "Yes" if policy.get("alerts") else "No",
            "Document Center": "Yes" if policy.get("documents") else "No",
            "Admin": "Yes" if policy.get("admin") else "No",
        })
    st.dataframe(pd.DataFrame(role_rows), use_container_width=True, hide_index=True)

    st.subheader("Data Files")
    rows = []
    for name, path in DATA_FILES.items():
        rows.append({
            "File": name,
            "Exists": path.exists(),
            "Size KB": round(path.stat().st_size / 1024, 1) if path.exists() else 0,
            "Modified": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S") if path.exists() else "",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.subheader("Backups")
    backups = sorted(BACKUP_DIR.glob("*"), reverse=True)
    if backups:
        st.dataframe(pd.DataFrame([{
            "Backup": b.name,
            "Size KB": round(b.stat().st_size / 1024, 1),
            "Modified": datetime.fromtimestamp(b.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        } for b in backups]), use_container_width=True, hide_index=True)
    else:
        st.info("No backups yet.")

    st.subheader("Google Drive Upload Configuration")
    st.code("""
[google_drive]
root_folder_id = "PASTE_LINK_CODES_FOLDER_ID"

[google_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "your-service-account@your-project.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
    """.strip())

    st.subheader("Email Configuration")
    st.code(
        """
[email]
smtp_host = "smtp.office365.com"
smtp_port = 587
smtp_user = "your-email@company.com"
smtp_password = "your-password-or-app-password"
sender = "your-email@company.com"
        """.strip()
    )



def render_session_bar() -> None:
    """Visible authenticated session bar shown above every page."""
    username = str(st.session_state.get("username", "") or "User")
    role = str(st.session_state.get("role", "viewer") or "viewer").lower()
    role_label = ROLE_DISPLAY_NAMES.get(role, role.title())
    last_login = str(st.session_state.get("last_login", "") or _format_login_time())

    st.markdown(
        f"""
        <style>
        .session-bar {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 14px;
            padding: 14px 18px;
            margin: 0 0 16px 0;
            border-radius: 18px;
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 55%, #0f766e 100%);
            color: #ffffff;
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.22);
            border: 1px solid rgba(255,255,255,0.20);
        }}
        .session-main {{
            display: flex;
            flex-wrap: wrap;
            gap: 10px 18px;
            align-items: center;
            font-weight: 800;
            font-size: 16px;
            line-height: 1.4;
        }}
        .session-pill {{
            display: inline-flex;
            align-items: center;
            gap: 7px;
            background: rgba(255,255,255,0.13);
            border: 1px solid rgba(255,255,255,0.20);
            padding: 8px 12px;
            border-radius: 999px;
            color: #f8fafc;
            white-space: nowrap;
        }}
        .session-pill strong {{
            color: #ffd400;
            font-weight: 900;
        }}
        .session-active {{
            color: #86efac;
            font-weight: 900;
        }}
        </style>
        <div class="session-bar">
            <div class="session-main">
                <span class="session-pill">👤 <strong>{username}</strong></span>
                <span class="session-pill">⚙️ Role: <strong>{role_label}</strong></span>
                <span class="session-pill">🕒 Last Login: <strong>{last_login}</strong></span>
                <span class="session-pill session-active">🔐 Session Active</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_spacer, col_logout = st.columns([8.7, 1.3])
    with col_logout:
        if st.button("🚪 Logout", use_container_width=True, key="top_logout"):
            _clear_login_query_params()
            st.session_state.clear()
            st.rerun()



def main() -> None:
    if not login_page():
        return

    role = st.session_state.get("role", "viewer")

    with st.sidebar:
        st.markdown("### Dawiyat PMO Portal V2")
        st.caption(f"User: {st.session_state.get('username','')}")

        pages = []
        if can("dashboard"):
            pages.append("Dashboard")
        if can("assistant"):
            pages.append("AI Executive Assistant")
        if can("alerts"):
            pages.append("Smart Alerts")
        if can("reports"):
            pages.append("Executive Reports")
            pages.append("📊 Executive PPT Builder")
        if can("upload"):
            pages.append("Upload CSV")
        if can("documents"):
            pages.append("📤 Document Upload Center")
        if can("admin"):
            pages.append("Admin Board")
        if not pages:
            pages = ["Dashboard"]

        # IMPORTANT: Streamlit reruns the script after every selectbox/file_uploader/button action.
        # Keep the selected page in session_state so choosing a Link Code inside
        # Document Upload Center never sends the user back to Dashboard.
        if st.session_state.get("force_dashboard") and "Dashboard" in pages:
            st.session_state["main_nav"] = "Dashboard"
            st.session_state["force_dashboard"] = False

        if st.session_state.get("force_document_upload_center") and "📤 Document Upload Center" in pages:
            st.session_state["main_nav"] = "📤 Document Upload Center"
            st.session_state["force_document_upload_center"] = False

        if st.session_state.get("force_ppt_builder") and "📊 Executive PPT Builder" in pages:
            st.session_state["main_nav"] = "📊 Executive PPT Builder"
            st.session_state["force_ppt_builder"] = False

        if st.session_state.get("force_dashboard") and "Dashboard" in pages:
            st.session_state["main_nav"] = "Dashboard"
            st.session_state["force_dashboard"] = False

        if st.session_state.get("main_nav") not in pages:
            st.session_state["main_nav"] = "Dashboard"

        page = st.radio(
            "Navigation",
            pages,
            key="main_nav",
            label_visibility="collapsed",
        )

    render_session_bar()

    if page == "Dashboard":
        render_dashboard()
    elif page == "AI Executive Assistant":
        ai_assistant_page()
    elif page == "Smart Alerts":
        smart_alerts_page()
    elif page == "Executive Reports":
        reports_page()
    elif page == "📊 Executive PPT Builder":
        executive_ppt_builder_page()
    elif page == "Upload CSV":
        upload_data_page()
    elif page == "📤 Document Upload Center":
        document_upload_page()
    elif page == "Admin Board":
        admin_page()


if __name__ == "__main__":
    main()
