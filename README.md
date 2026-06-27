# Dawiyat PMO Executive Portal V6.0 UI Professional Layout

Baseline: V5.9.9 Document Center UI Smart Filter Fixed

## Included in this build
- Professional left Sidebar V2 with PMO Portal branding.
- Dashboard pages grouped inside the sidebar.
- Smart Bulk Filter toggle moved into the sidebar with separator line.
- Governance / Quick Actions moved from the dashboard top area into the sidebar.
- Top Quick Actions block removed from the main dashboard canvas.
- Sidebar buttons remain permission-based.
- Existing Document Upload Center / Project Updates / Notification / Daily Digest / PPT Builder logic retained.

## Notes
- Persistence layer remains as in V5.9.9; Google Drive persistence should be implemented as the next backend phase.


## V6.0.3 Permission-Safe Sidebar Navigation Fix
- Preserves data/permissions.xlsx as the single source of truth.
- Moves dashboard tab navigation and governance actions to the Streamlit sidebar.
- Disables iframe/dashboard.html action navigation to avoid logout/return-to-start behavior.
- Keeps Smart Bulk Filter in the sidebar and renders the dashboard with the selected tab.
