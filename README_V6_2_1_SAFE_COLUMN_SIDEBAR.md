# Dawiyat PMO Executive Portal V6.2.1

Safe rebuild based on the last stable V5.9.9 codebase.

## Fixes
- Replaced unreliable embedded/HTML-only sidebar with a Streamlit functional left navigation column.
- Kept the requested sidebar structure: Dashboard Navigation, Smart Scope, Governance Actions, Logout.
- Removed Quick Actions from top dashboard page.
- Preserved permissions.xlsx, data files, formulas, filters and existing page functions.
- Dashboard tabs are driven by `st.session_state['dashboard_tab']` and opened in the same app.
- Hidden governance pages open through `st.session_state['active_hidden_page']`.
- Executive Reports iframe height increased and table internal scroll limits removed as much as possible.

## Notes
If the browser keeps an old state, use Streamlit Clear cache or hard refresh after deployment.
