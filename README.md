# Dawiyat PMO Executive Portal V5.9.3

Fixes included:
- Executive SOR Summary now classifies SOR Status = Created + 1st 50 Invoice Status = SOR not Create / Not Created / blank as Not Start, not Submitted.
- Submitted is counted only when 1st 50 Invoice Status is explicitly Submitted.
- Project Updates save now persists to project_updates.csv, change_log.csv, master operational snapshot, and local u_osp_work_order.csv for process reboot continuity.

Important: for permanent persistence across Streamlit Cloud redeploys, commit the updated data files/snapshots to GitHub or download them from Snapshot Center.
