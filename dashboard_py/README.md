# V4 Enterprise Architecture Notes

This build applies the first practical performance split:
- `u_osp_work_order.csv` is the main data source for operational and location fields.
- Streamlit CSV reads use cached loaders.
- The dashboard iframe renders only active/deferred sections to avoid browser freeze.

Next safe refactor step: move dashboard calculations from `dashboard.html` JavaScript to Python modules gradually.
