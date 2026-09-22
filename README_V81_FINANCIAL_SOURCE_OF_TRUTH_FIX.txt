V81 - FINANCIAL SOURCE-OF-TRUTH / INTERNAL ERROR FIX

- Payroll remains the source of truth for salary and statutory contributions.
- No 2nd-cutoff Approved/Paid payroll for a month means Contribution Shares is empty and all government pending amounts are zero.
- Legacy contribution rows with NULL/stale source_payroll_id are purged for the selected month.
- Auto-generated government expense rows are removed when their payroll source no longer exists.
- Financial Summary runs the production migration before ORM queries that depend on V80 columns.
- Existing payroll computation is unchanged.
- Added a Render/Gunicorn 500 logger so any remaining server exception appears in deployment logs.
