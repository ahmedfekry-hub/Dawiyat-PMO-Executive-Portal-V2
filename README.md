# Dawiyat PMO Executive Portal - Drive Upload Ready

## What is included
- Streamlit executive portal.
- Embedded HTML dashboard.
- Google Drive folder opening from Link Code Summary Table.
- Direct upload from Streamlit to Google Drive using Service Account.
- Option A folder structure per Link Code:
  - 01 Design
  - 02 Permit
  - 03 Photos
  - 04 PAT
  - 05 AsBuilt
  - 06 Handover
  - 07 Commercial

## Important security note
Never upload or commit the real Google Service Account JSON key to GitHub.
Place the values only in Streamlit Cloud > App settings > Secrets.
If the key was uploaded to GitHub or shared publicly, delete that key from Google Cloud and create a new one.

## Streamlit Secrets format
Use TOML format, not JSON:

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

## Google Drive setup
1. Enable Google Drive API in Google Cloud.
2. Create a Service Account.
3. Create a JSON key.
4. Share the root Drive folder `Dawiyat PMO Repository` with the Service Account client email as Editor.
5. Copy the root folder ID into `google_drive.root_folder_id`.

## How upload works
Open the Streamlit page `Document Upload`, choose a Link Code, then upload to Design / Permit / Photos / PAT / AsBuilt / Handover / Commercial. The app creates missing Link Code folders and subfolders automatically.
