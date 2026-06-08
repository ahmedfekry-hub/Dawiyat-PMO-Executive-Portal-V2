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
