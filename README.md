# Dawiyat PMO Executive Portal V2

Professional Streamlit Executive Portal with the same embedded HTML dashboard design and calculations.

## Included

- Direct dashboard opening, no HTML upload
- Login page with Dawiyat / MET branded style
- Admin / PMO / Viewer roles
- CSV upload from browser
- Backup previous CSV versions
- Arabic AI Executive Assistant
- Smart Alerts Dashboard
- PDF Executive Report download
- Email alerts framework using SMTP secrets
- Admin Board

## Main file path

```text
app.py
```

## Default Login

```text
admin / Admin@12345
pmo / PMO@12345
viewer / Viewer@12345
```

## Recommended Streamlit Secrets

```toml
[users.admin]
password = "AdminStrongPassword"
role = "admin"

[users.pmo]
password = "PMOStrongPassword"
role = "pmo"

[users.viewer]
password = "ViewerStrongPassword"
role = "viewer"

[email]
smtp_host = "smtp.office365.com"
smtp_port = 587
smtp_user = "your-email@company.com"
smtp_password = "your-password-or-app-password"
sender = "your-email@company.com"
```

## Permanent URL

In Streamlit Cloud, rename the app URL from app settings to a clean name such as:

```text
dawiyat-pmo
```

Then the app URL becomes close to:

```text
https://dawiyat-pmo.streamlit.app
```

subject to Streamlit availability.


## How to define permissions for each user

Permissions are controlled from Streamlit Secrets, not from the code.

Use one of these roles:

- `admin`: full access, upload CSV, backups, admin board, email alerts.
- `pmo`: dashboard, assistant, smart alerts, reports, CSV upload. No admin board and no email configuration.
- `viewer`: read-only access to dashboard, assistant, alerts, and reports. No upload and no admin board.

Example:

```toml
[users.ahmed]
password = "StrongPassword123"
role = "admin"

[users.pmo_team]
password = "PMO123"
role = "pmo"

[users.board]
password = "ViewOnly123"
role = "viewer"
```

In Streamlit Cloud:
App → Settings → Secrets → paste the TOML block → Save → Reboot app.


## Production Login Screen

The default first-login credentials panel has been hidden from the public login page. Manage accounts from Streamlit Secrets.


## Final Login Users

The public login page no longer shows default credentials.

Use these credentials unless you override them in Streamlit Cloud Secrets:

```text
Username: ahmedfekry
Password: 20020099
Role: admin

Username: pmo_team
Password: PMO12345
Role: pmo

Username: board
Password: Met_12345
Role: viewer
```

Recommended Streamlit Cloud Secrets:

```toml
[users.ahmedfekry]
password = "20020099"
role = "admin"

[users.pmo_team]
password = "PMO12345"
role = "pmo"

[users.board]
password = "Met_12345"
role = "viewer"
```

After saving Secrets, reboot the Streamlit app.

## Google Drive Direct Document Upload

This version adds a Streamlit page called **Document Upload** for Admin and PMO users.

What it does:
- Select a Link Code.
- Select document type: Design, Permit, Photos, PAT, AsBuilt, Handover, or Other.
- Upload one or more files directly to Google Drive.
- If the Link Code already has a `Document_Link`, files are uploaded under that folder.
- If `Document_Link` is empty, the app creates a Link Code folder under the configured root folder and updates `u_osp_work_order.csv` with the new folder link.
- Optional subfolders are created per document type.

Required Streamlit Secrets:

```toml
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
```

Important: Share the Google Drive root folder with the service account `client_email` as **Editor**.
