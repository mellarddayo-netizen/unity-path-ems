V80 - FINANCIAL SOURCE OF TRUTH FIX

1. Payroll is the source of truth for salary and statutory contribution obligations.
2. A contribution record is valid only when linked to an existing Approved/Paid 2nd-cutoff payroll for the same employee and month.
3. Legacy/orphan MonthlyContribution records are removed automatically when financial pages are opened.
4. Contribution Shares displays only valid payroll-linked contribution records.
5. Profit & Expenses Summary calculates pending government obligations only from valid payroll-linked contribution records.
6. Payroll salary expenses are removed automatically when their source payroll is deleted.
7. Payroll contribution expense rows are removed automatically when their source payroll-linked contribution records no longer exist, including previously Paid rows.
8. Added a production migration for source_payroll_id on monthly_contributions and source_type/source_key on expenses for older databases.
9. Existing payroll computation is not changed.
