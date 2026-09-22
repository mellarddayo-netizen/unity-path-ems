UNITY PATH EMS - V78 Financial Link Fixes

1. Payroll is automatically Approved when generated. Draft payroll is no longer used for the normal payroll workflow.
2. Existing payroll records with status Draft are converted to Approved at startup.
3. Approved payroll automatically creates a Pending Salary / Payroll expense.
4. Paid payroll automatically becomes a Paid Salary / Payroll expense.
5. Deleting a payroll removes its linked contribution record and stale salary/contribution expense records are automatically cleaned from financial reports.
6. Contribution expense tracker uses employer/company share as the company expense. Employee statutory deductions are not counted as additional company expense.
7. Pending government obligations in Profit & Expenses Summary show employee share due + employer share until the remittance expense is marked Paid.
8. Commission Excel upload is the source of company income: Total Collection goes directly to Total Income for the matching month.
9. Uploaded Agent Commission is automatically included as a company expense in Net Profit. The Commission Report itself has no Paid/Pending status.
10. Profit & Expenses Summary no longer requires a separate manual Income entry for normal company income.
11. Dashboard Total Income and Net Profit use Commission Report Total Collection / Commission as the income/commission source.
12. Clean package: old nested ZIP files were removed.

Important: existing database data is preserved. The startup migration only converts old payroll Draft records to Approved.
