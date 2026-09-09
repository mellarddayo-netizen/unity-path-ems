EMS V37 — PAYROLL / ATTENDANCE / STATUTORY FIXES

This version is based on V36 and fixes the payroll/contribution reconciliation issues found during testing.

KEY FIXES
1. INVALID OUT is never payable. Payroll calculations only count completed same-day Present records.
2. Legacy attendance rows are re-normalized on startup using the current attendance rules.
3. Invalid Out records cannot contribute to worked days, hours, late, undertime, overtime, allowance, incentive, or basic pay.
4. Final Pay uses the same valid-attendance rule.
5. SSS/Pag-IBIG/PhilHealth contribution records are linked to the 2nd-cutoff payroll and store the actual payroll deduction as Collected.
6. Contribution Shares now supports selecting a month/year and shows Employee Due, Actual Deducted, Uncollected, and Company Share.
7. The Contribution Shares page uses the finalized monthly contribution record when available, so it reconciles with the 2nd-cutoff payroll.
8. PhilHealth MBS uses fixed basic salary (daily rate × configured monthly working days) and excludes overtime, allowances, absences, tardiness, and undertime from the MBS.
9. SSS uses actual monthly remuneration and the current SSS MSC schedule (5% employee, 10% employer, employer-only EC).
10. Pag-IBIG uses monthly compensation capped at P5,000 for mandatory contribution computation.
11. BIR withholding tax is computed per semi-monthly payroll period, including the 1st cutoff. The 1st cutoff still has zero SSS/PhilHealth/Pag-IBIG employee deduction under the EMS policy.
12. Unpaid late/undertime reduce the taxable compensation base.
13. Paid payroll can still be deleted by Admin, and linked loan/contribution records are reversed/removed as implemented in V36.

PAYROLL RULES
- 1st payroll: day 1–15; no employee SSS/PhilHealth/Pag-IBIG deduction.
- 2nd payroll: day 16–last day; monthly SSS/PhilHealth/Pag-IBIG obligation is finalized/collected here.
- If 2nd payroll gross is insufficient, only the collectible amount is deducted and the remainder is shown as Uncollected.
- No Time In + No Time Out = Absent.
- Time In + no Time Out = Incomplete.
- Cross-date Time Out = Invalid Out.
- Invalid Out / Incomplete / Absent = 0 paid day until corrected.
- Regular schedule: 8:00 AM–5:00 PM, 12:00 PM–1:00 PM unpaid lunch.
- Early arrival does not offset undertime.
- OT starts after 5:00 PM and is paid only when approved by Admin.

IMPORTANT
Back up the existing SQLite database before replacing project files.
After copying V37 into the project folder, run:
    pip install -r requirements.txt
    python app.py

The environment used to audit this package did not have Flask installed and could not install it because external package access was unavailable. Python syntax compilation of app.py passed successfully.
