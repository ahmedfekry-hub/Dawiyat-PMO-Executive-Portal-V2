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

## V4.1 Smart Bulk Upload UI Fixed
- Added a visible Streamlit-native **Smart Bulk Filter** panel above the dashboard.
- Upload supports `.xlsx`, `.xls`, and `.csv` files containing any column similar to Link Code and/or Work Order.
- Added searchable chip-style `Scan Link Codes` multi-select.
- Added searchable chip-style `Scan Work Orders` multi-select.
- Added manual Work Order paste box for large lists.
- The active selection is injected into the dashboard iframe and applies to Dashboard + PMO Audit scope.

## V4.2 Smart Bulk Filter Hidden Panel + Upload Error Fix
- Smart Bulk Filter is now hidden by default and appears only after clicking **Show Smart Bulk Filter**.
- Added a small active-status summary and quick Clear button when a bulk filter is active.
- Fixed Streamlit multiselect error after uploading Excel/CSV with Link Codes or Work Orders that are not already in current dashboard options.
- Uploaded Excel/CSV values are now inserted into the Scan Link Codes / Scan Work Orders options before widgets render.


## V4.3 Executive Reports - WO Billing & Handover Status Report
- Added a new report in Executive Reports: WO Billing & Handover Status Report.
- Columns: Link Code, Work Order, Region, District, implementation update, SOR Status, First 50% status, Second 50% status, PAT Status, Handover O&M _Status, Handover Consultant _Status.
- Report is linked to all Dashboard filters, PMO-linked filtering, and Smart Bulk Filter.
- Includes Export Excel button.

## V4.4 Executive Reports WO Status Report Fix
- Added City column to WO Billing & Handover Status Report.
- Corrected implementation update logic to use the real `implementation update` column only.
- Prevented numeric fallback values such as 0/1 from appearing when the implementation update field is not the intended source.
- Enabled the report Export Excel button with a direct report-level export action.


## V4.5 WO Status Report Export Fix
- Replaced the table HTML export for WO Billing & Handover Status Report with a dedicated XLSX export function.
- Export now reads the same filtered dataset used by the report, including Smart Bulk Filter.
- Added CSV fallback if the XLSX browser library is unavailable.


## V4.6 WO Status Export Button Fix
- Fixed WO Billing & Handover Status Report Export Excel button by using the same generic table export engine used by the other dashboard tables.
- Added explicit delegated click binding for the report export button to avoid inline onclick/cache issues inside Streamlit iframe.
- Export now downloads the visible filtered report table as Excel-compatible `.xls`.


## V4.7 WO Status Export Click Fix
- Root cause fixed: CSS in Executive Reports was disabling pointer events for all report-page buttons.
- The CSS is now scoped only to copied snapshot buttons, so the WO Billing & Handover Status Report Export Excel button is clickable.
- Export function now uses the visible filtered table and has a fallback to build an export table from filtered rows.

## V4.8 Hidden Action Pages UI
- Hidden sidebar navigation for Document Upload Center, Executive PPT Builder, and Admin Board.
- These pages now open only from compact Dashboard Quick Actions buttons, according to user permissions.
- Back to Dashboard clears the hidden-page route safely without Streamlit session-state errors.
- Smart Bulk Filter remains hidden by default and opens only through its Show button.


## V4.9 Export PDF Permission Gate Fix
- Fixed Dashboard HTML global Export PDF Report button visibility.
- The button is now hidden server-side with CSS when User_Component_Access has Export PDF = No for the user.
- This is not a permissions.xlsx issue; it was a static HTML dashboard button that needed a hard permission gate before iframe rendering.


## V5.0 Dynamic Export PDF Permission Fix
- Top dashboard **Export PDF Report** is now controlled by a dedicated component row: `Global PDF Report`.
- Table-level `Export PDF` permissions no longer control or open the full-page PDF export.
- Admin Board edits to `User_Component_Access` control this behavior the same way as the other permissions after refresh/rerun.
- The included permissions.xlsx contains `Global PDF Report` for every user; set Show/Export PDF to Yes only for users allowed to export the full dashboard PDF.

## V5.2 Selective PDF Export Permissions
- `Global PDF Report` controls only the visibility of the top `Export PDF Report` button.
- The PDF content is now controlled by each row in `User_Component_Access` using `Export PDF = Yes/No`.
- If `Global PDF Report = Yes` but no component has `Export PDF = Yes`, the user sees an alert and no PDF is generated.
- If only `Link Code Summary Table` has `Export PDF = Yes`, the global PDF exports only that table, not the full dashboard.
- `Global PDF Report` itself is not counted as a table/component export permission.


## V5.3 Selective PDF Export Enforcement Fix
- Global PDF Report now controls only whether the main Export PDF Report button appears.
- The PDF content is generated into a separate print container and includes only components with Export PDF = Yes.
- If no components have Export PDF = Yes, the user receives an alert and no PDF export starts.
- Global PDF Report accepts either Show = Yes or Export PDF = Yes for the dedicated Global PDF Report row.
- The original dashboard full-page print listener is replaced by the permission-aware selective PDF handler.


## V5.4 Selective PDF Reports Page Component Matching Fix
- Selective PDF export now recognizes Executive Reports page cards such as Executive Portfolio Summary & Cost Exposure, Executive KPI Cards, SOR Summary & Revenue Exposure, Overall Stages Summary & Cost Exposure, and WO Billing & Handover Status Report.
- The PDF matching engine now includes .report-section-card blocks instead of only older .panel blocks.
- Enabled Export PDF rows in Admin Board can now export Executive Reports page components individually.


## V5.5 User Permission Conflict + KSA Time Fix
- Fixed duplicated component-name conflict: Show=Yes now wins when the same component appears in another page with Show=No.
- Tables & Exports remains visible when the user has Tables & Exports page permission and its table components are Show=Yes.
- Admin Board edits still write directly to data/permissions.xlsx; users see changes after browser refresh or Logout/Login.
- Last Login and generated timestamps now use Saudi Arabia time (UTC+3).
- Added best-effort CSS/config to hide Streamlit Manage App toolbar elements from the portal UI. Streamlit Cloud may still show Manage App to app owners/collaborators at the platform level.


## V5.6 Tables Export Excel Visibility Fix
- Fixed Tables & Exports buttons hidden while Export Excel permissions are Yes.
- Added explicit allowed_excel_components logic from User_Component_Access.
- Scoped component export hiding to real panels/cards only, preventing parent containers from hiding valid table export buttons.
- Added support for the generic Tables & Exports / Export Excel permission row to keep all enabled table export buttons visible.
- Included the latest uploaded permissions.xlsx.


## V5.6.1 Data Structure + Signature Fix
- Restored the required project folder structure: `data/`, `dashboard/`, `assets/`, `.streamlit/`.
- Dashboard now reads CSV files from `data/u_osp_work_order.csv`, `data/District.csv`, and `data/Penalties.csv`.
- Fixed the prepared-by signature in both `app.py` and `dashboard/dashboard.html` to: `Prepared by Eng/Ahmed Fekry - Quality & Performance Director (PMO)`.
- Kept the dashboard title as: `Dawiyat Executive Project Dashboard`.
