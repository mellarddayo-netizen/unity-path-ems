UNITY PATH EMS V79 – Financial Cleanup Fix

1. Payroll remains the source of truth for salary and statutory contribution obligations.
2. Payroll generation is Approved by default; there is no Draft payroll workflow.
3. Deleting a payroll removes its linked MonthlyContribution and salary expense records.
4. Contribution Shares no longer calculates statutory amounts for every active employee when no payroll exists. It displays only finalized monthly contribution records created by the 2nd payroll.
5. Profit & Expenses Summary uses those payroll-linked contribution records, so deleting the payroll removes the corresponding pending amounts.
6. Employee share is a remittance liability; employer/company share is the company expense.
7. The Contribution Shares empty-state message now clearly says no payroll contribution has been generated for the selected month.
