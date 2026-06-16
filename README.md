# Dawiyat PMO Executive Portal V2.5 - Google Drive Upload Center

This version includes a professional Streamlit Document Upload Center integrated with Google Drive.

## What is included

- Role-based login: admin / pmo / viewer
- Executive HTML dashboard
- CSV upload with backup
- Smart Alerts
- Arabic AI Executive Assistant
- PDF Executive Report
- Google Drive Document Upload Center

## Document Upload Center

The sidebar page **Document Upload Center** allows Admin and PMO users to upload documents directly to Google Drive for each Link Code.

Standard folder structure per Link Code:

```text
JED-HRR2-RIYA-60
├── 01 Design
├── 02 Permit
├── 03 Photos
├── 04 PAT
├── 05 AsBuilt
├── 06 Handover
└── 07 Commercial
```

## Streamlit Secrets

Do not upload Google JSON keys to GitHub. Put credentials only in Streamlit Cloud > App settings > Secrets.

```toml
[google_drive]
root_folder_id = "PASTE_DAWIYAT_PMO_REPOSITORY_FOLDER_ID"

[google_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = """-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----
"""
client_email = "your-service-account@your-project.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

The app also supports `[gcp_service_account]` if you already used that name.

## Required Google Drive setup

1. Enable Google Drive API in Google Cloud.
2. Create a Service Account.
3. Create a JSON key, but do not upload it to GitHub.
4. Copy its values into Streamlit Secrets as TOML.
5. Share the root Google Drive folder with the service account email as **Editor**.

## Security

`.gitignore` blocks:

- `*.json`
- `.streamlit/secrets.toml`
- `__pycache__/`
- `*.pyc`



## أين يظهر Document Upload Center؟
بعد تسجيل الدخول، افتح القائمة الجانبية في Streamlit واختر:

`📤 Document Upload Center`

ملاحظة: رفع الملفات لا يظهر داخل HTML Dashboard نفسه لأنه يعتمد على Streamlit native widgets و Google Drive API. لذلك تم وضعه كصفحة مستقلة داخل Streamlit مع زر واضح أعلى صفحة Dashboard للانتقال إليه.


## V2.7 Fixes
- Added **Document Upload Center — Status Preview** between Link Code Summary Table and Civil/Fiber Completion Summary inside the HTML dashboard.
- Fixed Google Service Account private key normalization for Streamlit TOML secrets.
- Upload Center remains available as a native Streamlit page: **📤 Document Upload Center**.


## Document Upload Center - Required Secrets

To allow the app to create new Link Code folders and upload files, add this in Streamlit Cloud > App settings > Secrets:

```toml
[google_drive]
root_folder_id = "PASTE_THE_GOOGLE_DRIVE_LINK_CODES_FOLDER_ID_HERE"

[google_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = """-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----
"""
client_email = "dawiyat-pmo-drive@dawiyat-pmo-portal.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

Important:
- `root_folder_id` should be the Google Drive folder ID of the **Link Codes** folder.
- Share that **Link Codes** folder with the Service Account email as **Editor**.
- Do not upload JSON key files to GitHub.

## V2.11 Manual Google Drive Upload Mode

This version disables direct Streamlit-to-Google-Drive file upload for normal My Drive accounts to avoid the Google Service Account storage quota limitation.

Workflow:
1. Open **📂 Document Upload Center**.
2. Select a Link Code.
3. Click **Open Link Code Folder** or **Open Subfolder**.
4. Upload files manually inside Google Drive under the correct subfolder:
   - 01 Design
   - 02 Permit
   - 03 Photos
   - 04 PAT
   - 05 AsBuilt
   - 06 Handover
   - 07 Commercial
5. Return to Streamlit and click **Refresh Document Status**.

Required Streamlit Secrets still include Google Drive service account credentials and:

```toml
[google_drive]
root_folder_id = "YOUR_LINK_CODES_FOLDER_ID"
```

The service account only scans folders/files; it does not upload files in this mode.

## Role-Based Access Control V2.14

Allowed roles: `admin`, `pmo`, `board`, `finance`, `viewer`.

| Role | Dashboard Pages | Upload CSV | Export | Smart Alerts | Document Center | Admin Board |
|---|---|---:|---|---|---:|---:|
| Admin | All tabs | Yes | Excel/PDF | Yes | Yes | Yes |
| PMO | All tabs except Admin Board | Yes | Excel/PDF | Yes | Yes | No |
| Board | Executive Overview + Executive PMO Report Assistant | No | PDF only | No | No | No |
| Finance | Tables & Exports + Executive PMO Report Assistant | No | Excel/PDF | Yes | No | No |
| Viewer | Executive Overview only | No | No | No | No | No |

Update users from Streamlit Secrets only. Do not edit `app.py` for every user change.

```toml
[users.finance]
password = "Finance12345"
role = "finance"
```


## V2.19 Session / Logout / Refresh Update
- Added visible Logout button in the sidebar.
- Added Refresh button that clears cached data and refreshes the current page without logging the user out.
- Removed Role display from the sidebar; only username is shown.
- Added signed browser-refresh session persistence so pressing the browser refresh button does not force a new login.
- Logout clears the persisted session parameters and returns to the login page.

Optional Streamlit Secrets:
```toml
[session]
secret = "PUT_A_LONG_RANDOM_SECRET_HERE"
```
If omitted, the app uses a deterministic fallback. For production, add a long random session secret.


## V41 Enterprise Permission Management

Supports page, dashboard tab, export, button, and table visibility control from Streamlit Secrets.

Example:

```toml
[users.tamer_solyman]
password = "Tamer@12345$"
role = "finance"

[users.board_member_1]
password = "Board@12345"
role = "board"
pages = ["Dashboard", "Executive Reports", "📊 Executive PPT Builder"]
dashboard_tabs = ["overview", "performance", "reports"]
export_excel = false
export_pdf = true
export_ppt = true
hide_buttons = ["Export Excel", "Upload", "Delete"]
hide_tables = ["PMO Audit", "Missing MET Actual", "Raw Data"]
```

After changing Secrets, save and reboot the app.



## V42 Excel Permission Engine

This version can read permissions from:

```text
data/permissions.xlsx
```

### Supported sheets

- `Users`: Username, Password, Role, Active
- `Role_Page_Access`: page visibility by role
- `Role_Component_Access`: table/component visibility and export permissions by role
- `User_Override`: optional user-specific exceptions
- `Reference_Lists`, `How_To_Use`: helper sheets

### How to update permissions

1. Edit `data/permissions.xlsx`.
2. Set cells to `Yes` or `No`.
3. Upload the workbook to GitHub inside the `data` folder, or upload it from the Admin Board.
4. Reboot the Streamlit app.

Secrets are still supported and can override Excel permissions when needed.

## V42.3 Strict RBAC Rendering Fix

- Disallowed dashboard tabs are now hidden with hard CSS and repeated JS enforcement.
- Component/table matching now ignores the word Executive and punctuation differences, so `Executive SOR Summary` also matches dashboard title `SOR Summary`.
- MutationObserver keeps permissions applied after dashboard re-rendering.
- If a page is visible but all components inside it are set to `No`, the page may show only filters. Set the required components to `Yes` in `Role_Component_Access` or `User_Override`.


## V43 User-Based Permission Engine

V43 removes role-based permission decisions. The `Role` column in `Users` is now only a display label / department.

The app reads permissions directly from:

- `Users`
- `User_Page_Access`
- `User_Component_Access`

Legacy sheets such as `Role_Page_Access`, `Role_Component_Access`, and `User_Override` are no longer required for day-to-day permission control.

After changing `data/permissions.xlsx`, reboot the Streamlit app.


## V43.2 User-Based Permission Final Fix

Fixes:
- Department/Display Role is no longer forced to Viewer when it is not a legacy role key.
- Dashboard tab hiding now also targets `report-tab`, so Executive Reports will disappear when not allowed.
- No automatic fallback to Dashboard when a username has no assigned pages.
- Latest `data/permissions.xlsx` is included in the package.


## V43.3 Permission Auto-Refresh Engine

This edition fixes stale permission behavior after `data/permissions.xlsx` changes.

Key points:
- Permissions are username-based only.
- `Role` / `Department` is display-only.
- The app reads `permissions.xlsx` from disk on every rerun and tracks file signature.
- If a user is disabled or removed while logged in, the app forces logout on the next rerun.
- Admin Board includes `Reload Permissions`, current file signature, and last modified timestamp.
- `User_Component_Access` has been expanded to include all known dashboard tables/sections from `dashboard.html`.
- Default-deny logic is applied to components: anything listed in the workbook but not explicitly `Show = Yes` for that username is hidden.

Recommended operating process:
1. Edit `data/permissions.xlsx`.
2. Upload/commit to GitHub.
3. Wait for Streamlit Cloud to redeploy, or click Reboot App.
4. Users should Logout/Login after page-level permission changes.


## V44 Admin Board + User-Based Permissions Only

This build applies the requested permission model:

- Added **Admin Board** page visible only to username `ahmedfekry`.
- Admin Board includes three live sections from `data/permissions.xlsx`:
  - **Active Users**
  - **Page Access**
  - **Component Access**
- Removed role-based permission logic from the effective permission engine. Department/Role is display only.
- Effective access is now controlled only by:
  - `Users`
  - `User_Page_Access`
  - `User_Component_Access`
- Added permission auto-refresh behavior: the app reads `permissions.xlsx` on rerun and tracks file signature/modified time.
- Improved loading performance by keeping heavy pages/data isolated by navigation and caching the dashboard HTML by modified time.

### Important
Only `ahmedfekry` should have `Admin Board = Yes` in `User_Page_Access`. Other users must remain `No`.

## V3 Admin Board UI Fix 2
- Added visible **← Back to Dashboard** button at the top of Admin Board.
- Added **Filter Admin Board by User** dropdown.
- The user filter applies to Active Users, Page Access, and Component Access tables.


## V3.4 Permission Fix
- permissions.xlsx is now the only authority for page access, dashboard tabs, components, and export buttons.
- Old page/component values inside Streamlit Secrets are ignored, so GitHub updates to data/permissions.xlsx take effect after Streamlit redeploy/reboot.
- Active=Yes/No remains controlled from the Users sheet.
- Dashboard tab CSS injection fixed so hidden tabs disappear immediately.


## V3.5 Dashboard / PMO Audit Filter Linkage Fix
- Global Dashboard Region / City / District filters now cascade correctly using `data/District.csv`.
- PMO Audit `Updated Region`, `Updated City`, and `WO Districts` filters are linked with the main dashboard filters.
- Changing Region clears City and District automatically.
- Changing City clears District automatically.
- The `Region` column in `u_osp_work_order.csv` remains ignored for location filtering; District.csv is the source of truth.

## V3.6 Live Admin Permission Editor
- Admin Board now supports direct inline editing using Streamlit `st.data_editor`.
- Ahmed can update Users, Page Access, and Component Access directly from Admin Board without uploading `permissions.xlsx` manually.
- Save buttons write changes back to `data/permissions.xlsx` on the running app.
- Active user sessions auto-refresh every 12 seconds outside Admin Board, so users receive permission changes quickly without manual logout/login.
- GitHub `permissions.xlsx` remains supported as the baseline file after redeploy, while Admin Board edits apply immediately to the current running deployment.

## V3.7 Manual Permission Refresh Fix
- Removed the timed 12-second permission auto-refresh to keep user pages stable.
- Admin Board can still edit and save permissions directly to `data/permissions.xlsx`.
- Other users will receive permission changes only after browser refresh, Streamlit rerun, or Logout/Login.
- This prevents active dashboard pages from resetting while users are working.


## V3.8 PPT Builder Page Gate Fix
- Executive PPT Builder page/button now depends only on User_Page_Access = Yes.
- Export PPT permission inside User_Component_Access no longer opens or grants the PPT Builder page.
- Added server-side guard: users without ppt_builder page permission cannot open it even via session/navigation.


## V3.9 Admin Back + Document Dates Fix
- Fixed Admin Board Back to Dashboard Streamlit session-state error.
- Removed repeated manual-upload info message from every document stage card.
- Added Uploaded / Created Date and Modified Date for all document stages 01 Design through 07 Commercial in Document Upload Center status scanning and export.

## V4.0 Smart Bulk Site Filter + Search Tokens
- Added **Smart Bulk Filter** inside the HTML dashboard.
- Upload Excel/CSV with any column similar to **Link Code** or **Work Order**; the dashboard detects the columns automatically.
- Added token-style search/add filter for **Link Code** similar to Scan Link Codes.
- Added token-style search/add filter for **Work Order**.
- Values can be pasted in bulk separated by comma, semicolon, tab, or new line.
- **Reset Filters** clears Link Code tokens, Work Order tokens, and any uploaded site list filter across the dashboard and PMO Audit.
