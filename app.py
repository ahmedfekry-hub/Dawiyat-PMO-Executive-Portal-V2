
import base64
import hashlib
import io
import json
import re
import shutil
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Dict, List, Tuple

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
DISTRICT_PATH = DATA_DIR / "District.csv"

DATA_FILES = {
    "u_osp_work_order.csv": WO_PATH,
    "Penalties.csv": PENALTIES_PATH,
    "District.csv": DISTRICT_PATH,
}

ASSETS_DIR = BASE_DIR / "assets"
DAWIYAT_LOGO_PATH = ASSETS_DIR / "dawiyat_logo.jpg"
MET_LOGO_PATH = ASSETS_DIR / "met_logo.jpg"

DOCUMENT_TYPES = ["Design", "Permit", "Photos", "PAT", "AsBuilt", "Handover", "Other"]
GOOGLE_DRIVE_FOLDER_MIME = "application/vnd.google-apps.folder"


def image_to_base64(path: Path) -> str:
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode("ascii")



ROLE_PERMISSIONS = {
    "admin": {
        "dashboard": True,
        "assistant": True,
        "alerts": True,
        "reports": True,
        "admin": True,
        "upload": True,
        "email": True,
        "documents": True,
    },
    "pmo": {
        "dashboard": True,
        "assistant": True,
        "alerts": True,
        "reports": True,
        "admin": False,
        "upload": True,
        "email": False,
        "documents": True,
    },
    "viewer": {
        "dashboard": True,
        "assistant": True,
        "alerts": True,
        "reports": True,
        "admin": False,
        "upload": False,
        "email": False,
        "documents": False,
    },
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
    max-width: 880px;
    margin: 5vh auto 2vh;
    padding: 0;
    border-radius: 34px;
    background:
        radial-gradient(circle at top left, rgba(20,184,166,.18), transparent 32%),
        radial-gradient(circle at top right, rgba(245,158,11,.18), transparent 34%),
        linear-gradient(135deg,#ffffff 0%,#f8fbff 100%);
    border: 1px solid #d9e3ef;
    box-shadow: 0 28px 70px rgba(15,23,42,.16);
    overflow: hidden;
    text-align: center;
}
.portal-login-top {
    height: 8px;
    background: linear-gradient(90deg,#f97316,#14b8a6,#2563eb);
}
.portal-login-inner {
    padding: 38px 46px 42px;
}
.logo-row {
    display:grid;
    grid-template-columns: 1.25fr .85fr;
    gap:22px;
    align-items:center;
    margin-bottom:26px;
}
.logo-card {
    height:92px;
    border:1px solid #e6edf5;
    border-radius:24px;
    display:flex;
    align-items:center;
    justify-content:center;
    background:#fff;
    box-shadow: 0 12px 28px rgba(15,23,42,.08);
    padding:14px 20px;
}
.logo-card img {
    max-width:100%;
    max-height:76px;
    object-fit:contain;
}
.logo-card.met-logo img {
    max-height:72px;
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
    margin-bottom:14px;
}
.portal-title {
    font-size: 40px;
    font-weight: 1000;
    color: #10223a;
    margin-bottom: 12px;
    letter-spacing:-.04em;
}
.portal-subtitle {
    color: #526985;
    font-size: 17px;
    line-height: 1.8;
    max-width:720px;
    margin:0 auto;
}
.portal-feature-row {
    display:grid;
    grid-template-columns: repeat(3,1fr);
    gap:12px;
    margin-top:26px;
}
.portal-feature {
    border:1px solid #e6edf5;
    border-radius:18px;
    background:rgba(255,255,255,.75);
    padding:13px 14px;
    color:#10223a;
    font-size:13px;
    font-weight:800;
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


def get_users() -> Dict[str, Dict[str, str]]:
    """
    Recommended Streamlit secrets format:
    [users.admin]
    password = "AdminStrongPassword"
    role = "admin"

    [users.pmo]
    password = "PMOStrongPassword"
    role = "pmo"

    [users.viewer]
    password = "ViewerStrongPassword"
    role = "viewer"
    """
    try:
        raw = dict(st.secrets.get("users", {}))
        if raw:
            users = {}
            for username, data in raw.items():
                if isinstance(data, dict):
                    users[str(username)] = {
                        "password": str(data.get("password", "")),
                        "role": str(data.get("role", "viewer")).lower(),
                    }
                else:
                    users[str(username)] = {"password": str(data), "role": "viewer"}
            return users
    except Exception:
        pass

    return {
        "ahmedfekry": {"password": "20020099", "role": "admin"},
        "pmo_team": {"password": "PMO12345", "role": "pmo"},
        "board": {"password": "Met_12345", "role": "viewer"},
    }


def can(permission: str) -> bool:
    role = st.session_state.get("role", "viewer")
    return ROLE_PERMISSIONS.get(role, ROLE_PERMISSIONS["viewer"]).get(permission, False)


def login_page() -> bool:
    if st.session_state.get("authenticated"):
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

    c1, c2, c3 = st.columns([1, 1.15, 1])
    with c2:
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        login = st.button("Login", use_container_width=True)

        if login:
            users = get_users()
            if username in users and password == users[username]["password"]:
                role = users[username].get("role", "viewer").lower()
                if role not in ROLE_PERMISSIONS:
                    role = "viewer"
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.session_state["role"] = role
                st.rerun()
            else:
                st.error("Invalid username or password.")

        st.caption("For account access, please contact the PMO System Administrator.")

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


def build_initial_raw() -> Dict[str, List[dict]]:
    return {
        "workorders": df_to_records(safe_read_csv(WO_PATH)),
        "penalties": df_to_records(safe_read_csv(PENALTIES_PATH)),
        "districts": df_to_records(safe_read_csv(DISTRICT_PATH)),
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

    portal_patch = """
<style>
.file-label, #apply-imports { display: none !important; }
.header-actions::after {
    content: "Data linked directly from Version 2 Executive Portal";
    display: block;
    color: #64748b;
    font-size: 12px;
    text-align: right;
    margin-top: 4px;
}
</style>
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
    """Create a Google Drive API service using Streamlit Secrets."""
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except Exception as exc:
        raise RuntimeError("Google Drive libraries are not installed. Install requirements.txt first.") from exc
    try:
        info = dict(st.secrets["google_service_account"])
    except Exception as exc:
        raise RuntimeError("Google service account secrets are not configured.") from exc
    if "private_key" in info:
        info["private_key"] = str(info["private_key"]).replace("\\n", "\n")
    scopes = ["https://www.googleapis.com/auth/drive"]
    credentials = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def get_drive_root_folder_id() -> str:
    try:
        return str(st.secrets["google_drive"]["root_folder_id"]).strip()
    except Exception:
        return ""


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
    folder_id = extract_drive_folder_id(existing_link)
    if folder_id:
        return folder_id, drive_folder_url(folder_id)
    root_id = get_drive_root_folder_id()
    if not root_id:
        raise RuntimeError("google_drive.root_folder_id is not configured in Streamlit Secrets.")
    folder_id = get_or_create_drive_folder(service, root_id, link_code)
    return folder_id, drive_folder_url(folder_id)


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


def document_upload_page() -> None:
    st.title("📤 Document Upload to Google Drive")
    st.caption("Upload files directly to the correct Google Drive folder for each Link Code.")

    if not can("documents"):
        st.error("You do not have permission to upload documents.")
        return

    wo = safe_read_csv(WO_PATH)
    if wo.empty:
        st.warning("u_osp_work_order.csv is missing or empty.")
        return
    if "Link Code" not in wo.columns:
        st.error("Link Code column is missing from u_osp_work_order.csv.")
        return

    links = sorted([x for x in wo["Link Code"].astype(str).str.strip().unique() if x])
    if not links:
        st.warning("No Link Codes found.")
        return

    with st.container(border=True):
        st.subheader("Upload Document")
        c1, c2 = st.columns([1.2, .8])
        with c1:
            link_code = st.selectbox("Link Code", links, index=0)
        with c2:
            doc_type = st.selectbox("Document Type", DOCUMENT_TYPES, index=0)

        current_folder_url = get_document_link_for_link(wo, link_code)
        if current_folder_url:
            st.success(f"Current folder linked for {link_code}")
            st.link_button("Open Current Folder", current_folder_url, use_container_width=True)
        else:
            st.warning("No Document_Link found for this Link Code. The app will create a folder under the configured root folder after upload.")

        uploaded_files = st.file_uploader("Select one or more files", accept_multiple_files=True, key="drive_doc_uploader")
        create_type_subfolder = st.checkbox("Upload inside document-type subfolder", value=True)

        if st.button("Upload to Google Drive", type="primary", use_container_width=True):
            if not uploaded_files:
                st.error("Please select at least one file.")
                return
            try:
                service = get_drive_service()
                link_folder_id, link_folder_url = ensure_link_folder(service, link_code, current_folder_url)
                target_folder_id = link_folder_id
                if create_type_subfolder:
                    target_folder_id = get_or_create_drive_folder(service, link_folder_id, doc_type)

                uploaded_rows = []
                for f in uploaded_files:
                    safe_name = f"{link_code}_{doc_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{f.name}"
                    file_url = upload_file_to_drive(service, target_folder_id, f, safe_name, f.type or "application/octet-stream")
                    uploaded_rows.append({"Link Code": link_code, "Document Type": doc_type, "File": f.name, "Drive Link": file_url})

                if not current_folder_url:
                    update_document_link_in_csv(link_code, link_folder_url)

                st.success(f"Uploaded {len(uploaded_rows)} file(s) successfully.")
                st.dataframe(pd.DataFrame(uploaded_rows), use_container_width=True, hide_index=True)
                st.link_button("Open Link Code Folder", link_folder_url, use_container_width=True)
            except Exception as exc:
                st.error(str(exc))
                st.info("Check Streamlit Secrets and ensure the Google Drive root folder is shared with the service account email as Editor.")

    with st.expander("Required Streamlit Secrets", expanded=False):
        st.code("""
[google_drive]
root_folder_id = "PASTE_DAWIYAT_PMO_REPOSITORY_FOLDER_ID"

[google_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "your-service-account@your-project.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
        """.strip())
        st.markdown("Share the Google Drive root folder with the service account email as **Editor**.")


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
    st.write("Current role:", st.session_state.get("role"))

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
root_folder_id = "PASTE_DAWIYAT_PMO_REPOSITORY_FOLDER_ID"

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


def main() -> None:
    if not login_page():
        return

    role = st.session_state.get("role", "viewer")

    with st.sidebar:
        st.markdown("### Dawiyat PMO Portal V2")
        st.caption(f"User: {st.session_state.get('username','')}")
        st.caption(f"Role: {role.upper()}")

        pages = ["Dashboard", "AI Executive Assistant", "Smart Alerts", "Executive Reports"]
        if can("upload"):
            pages.append("Upload CSV")
        if can("documents"):
            pages.append("Document Upload")
        if can("admin"):
            pages.append("Admin Board")

        page = st.radio("Navigation", pages, label_visibility="collapsed")

        if st.button("Logout", use_container_width=True):
            st.session_state.clear()
            st.rerun()

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
    elif page == "Document Upload":
        document_upload_page()
    elif page == "Admin Board":
        admin_page()


if __name__ == "__main__":
    main()
