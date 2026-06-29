# Dawiyat PMO Executive Portal V6.1 Enterprise Stable

This build restores a stable architecture after the embedded HTML sidebar routing issues.

## Key fixes
- Single functional navigation source: native Streamlit sidebar.
- Dashboard tabs controlled from the sidebar and injected into `dashboard.html`.
- Governance actions open Streamlit pages through `st.session_state`, not broken iframe query links.
- `permissions.xlsx` remains the source of page/component access.
- Quick Actions block is not rendered at the top of the dashboard.
- Embedded dashboard sidebar/drawer is disabled to avoid duplicate or non-functional buttons.
- Executive Reports receives a larger render height and table wrappers are forced visible to reduce missing lower sections.

## Preserved
- Existing data files.
- Existing dashboard calculations and filters.
- Project Updates Center.
- Notification Center.
- Document Upload Center.
- Executive PPT Builder.
- Admin Board permission logic.
