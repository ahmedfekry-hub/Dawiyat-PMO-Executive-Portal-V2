# Dawiyat PMO Executive Portal V5.9.2

Patch: Project Updates SOR Status alignment.

When a user updates SOR Status to Created but the old 1st 50 Invoice Status still contains `SOR not Create`, the dashboard now treats the invoice stage as `Not Start` for SOR/Billing calculations instead of incorrectly keeping it under SOR Not Create. This keeps Executive SOR Summary, WO Billing & Handover, filters, and cross-user updates aligned after refresh/login.
