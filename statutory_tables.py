"""Philippine statutory contribution tables used as the EMS payroll source of truth.

SSS: Circular No. 2024-006, effective January 2025.
PhilHealth: CY 2025 premium schedule (5%, P10,000 floor, P100,000 ceiling).
Pag-IBIG: HDMF Circular No. 274 contribution rates (max compensation P5,000).
"""

# SSS 2025 Business Employers and Employees.
# Each row is: compensation minimum, compensation maximum (None = no upper limit), MSC.
# The official schedule uses: below P5,250 -> P5,000 MSC; then P500 MSC steps.
SSS_2025_RANGES = []
SSS_2025_RANGES.append({"min": 0.0, "max": 5249.999999999, "msc": 5000.0})
for msc in range(5500, 35001, 500):
    minimum = float(msc - 250)
    maximum = None if msc == 35000 else float(msc + 249.999999999)
    SSS_2025_RANGES.append({"min": minimum, "max": maximum, "msc": float(msc)})


def sss_msc_from_remuneration(remuneration):
    """Return the exact 2025 SSS MSC bracket for an employed member."""
    compensation = max(0.0, float(remuneration or 0))
    if compensation <= 0:
        return 0.0
    # The published schedule is a set of compensation ranges, not ordinary
    # rounding. For every 2-decimal payroll amount, find the row whose lower
    # boundary has been reached. The final row has no upper limit.
    for row in reversed(SSS_2025_RANGES):
        if compensation >= row["min"]:
            return row["msc"]
    return 5000.0


def sss_contribution_row(msc):
    """Return the official SSS contribution components for an MSC."""
    msc = float(msc or 0)
    if msc <= 0:
        return {
            "regular_ss_er": 0.0, "mpf_er": 0.0, "ec": 0.0,
            "employer_total": 0.0, "regular_ss_ee": 0.0,
            "mpf_ee": 0.0, "employee_total": 0.0,
        }
    regular_ss_er = round(msc * 0.10, 2)
    regular_ss_ee = round(msc * 0.05, 2)
    mpf = max(0.0, msc - 20000.0)
    mpf_er = round(mpf * 0.10, 2)
    mpf_ee = round(mpf * 0.05, 2)
    ec = 10.0 if msc <= 14500.0 else 30.0
    return {
        "regular_ss_er": regular_ss_er,
        "mpf_er": mpf_er,
        "ec": ec,
        "employer_total": round(regular_ss_er + mpf_er + ec, 2),
        "regular_ss_ee": regular_ss_ee,
        "mpf_ee": mpf_ee,
        "employee_total": round(regular_ss_ee + mpf_ee, 2),
    }


SSS_2025_TABLE = []
for row in SSS_2025_RANGES:
    parts = sss_contribution_row(row["msc"])
    if row["max"] is None:
        range_label = f"P{row['min']:,.0f} and over"
    elif row["min"] == 0:
        range_label = f"Below P5,250"
    else:
        range_label = f"P{row['min']:,.0f} - P{row['max']:,.2f}"
    SSS_2025_TABLE.append({
        "range": range_label,
        "msc": row["msc"],
        **parts,
    })


PHILHEALTH_2025_TABLE = [
    {
        "monthly_basic_salary": "P10,000.00",
        "premium_rate": "5.0%",
        "monthly_premium": "P500.00",
        "employee_share": "P250.00",
        "employer_share": "P250.00",
    },
    {
        "monthly_basic_salary": "P10,000.01 - P99,999.99",
        "premium_rate": "5.0%",
        "monthly_premium": "P500.00 - P5,000.00",
        "employee_share": "50% of premium",
        "employer_share": "50% of premium",
    },
    {
        "monthly_basic_salary": "P100,000.00",
        "premium_rate": "5.0%",
        "monthly_premium": "P5,000.00",
        "employee_share": "P2,500.00",
        "employer_share": "P2,500.00",
    },
]

PAGIBIG_TABLE = [
    {
        "monthly_compensation": "P1,500.00 and below",
        "employee_rate": "1.0%",
        "employer_rate": "2.0%",
    },
    {
        "monthly_compensation": "Over P1,500.00",
        "employee_rate": "2.0%",
        "employer_rate": "2.0%",
    },
]
PAGIBIG_MAX_COMPENSATION = 5000.0
