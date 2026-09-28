V87 - STATUTORY TABLE SOURCE OF TRUTH

Payroll Settings now displays the statutory tables used by the payroll engine:

1. SSS 2025 Business Employers and Employees table (Circular No. 2024-006)
   - Exact compensation range -> MSC mapping
   - Regular SS, MPF, EC, employer total and employee total
   - The payroll engine uses the exact table ranges; it does NOT ordinary-round compensation to the next P500.

2. PhilHealth CY 2025 premium table
   - 5.0% premium rate
   - P10,000 floor / P100,000 ceiling
   - 50/50 employee and employer share
   - Monthly Basic Salary exclusions are preserved in the payroll calculation.

3. Pag-IBIG contribution table
   - P1,500 and below: 1% employee / 2% employer
   - Over P1,500: 2% employee / 2% employer
   - Maximum monthly compensation basis: P5,000

The payroll engine imports these tables from statutory_tables.py. No database migration is required for this change.

Validation performed:
- Python compile check passed for app.py, statutory_tables.py and models/payroll.py.
- SSS boundary checks were run, including P15,507.75 -> MSC P15,500 -> employee SSS P775.
- Full Flask runtime import could not be executed in this build environment because Flask is not installed there.
