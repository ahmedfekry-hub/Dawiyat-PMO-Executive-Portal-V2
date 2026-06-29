# V6.1.1 Safe Sidebar Navigation

This package restores the last stable V5.9.9 baseline and applies a safe Streamlit-native sidebar.

What changed:
- Quick Actions removed from the dashboard body.
- Smart Bulk Filter moved to the sidebar.
- Governance Actions moved to the sidebar.
- Logout moved to the sidebar.
- Existing permissions.xlsx, formulas, filters, dashboard.html, and data files are preserved.
- Modular folders were added as a safe scaffold for future refactoring without breaking the current production app.

Important: The actual runtime remains in app.py to avoid breaking the mature legacy logic. Full page extraction should be done gradually after this version passes UAT.
