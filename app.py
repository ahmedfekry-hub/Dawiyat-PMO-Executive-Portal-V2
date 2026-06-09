
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
    display:inline-flex;
    align-items:center;
    justify-content:center;
    width: min(860px, 92vw);
    margin: 0 auto;
    padding:16px 30px;
    border-radius:999px;
    background:linear-gradient(135deg,#111827,#1e3a8a);
    color:#facc15 !important;
    font-size:22px;
    line-height:1.35;
    font-weight:1000;
    box-shadow:0 14px 28px rgba(15,23,42,.24);
    border:1px solid rgba(255,255,255,.18);
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
            '<div class="account-access-row"><div class="account-access-pill">For account access, please contact Eng./Ahmed Fekry 
            (PMO System Administrator).</div></div>',
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
    elif page == "Upload CSV":
        upload_data_page()
    elif page == "📤 Document Upload Center":
        document_upload_page()
    elif page == "Admin Board":
        admin_page()


if __name__ == "__main__":
    main()
