from datetime import datetime
from extensions import db


class PayrollSettings(db.Model):
    __tablename__ = "payroll_settings"

    id = db.Column(db.Integer, primary_key=True)
    working_days_per_month = db.Column(db.Float, default=26)
    hours_per_day = db.Column(db.Float, default=8)
    overtime_multiplier = db.Column(db.Float, default=1.25)

    # Legacy fields retained for database compatibility. Statutory deductions
    # are now computed automatically in app.py and these values are not used.
    sss_employee_amount = db.Column(db.Float, default=0)
    philhealth_rate = db.Column(db.Float, default=0)
    pagibig_rate = db.Column(db.Float, default=0)
    pagibig_max_employee_contribution = db.Column(db.Float, default=0)
    withholding_tax_rate = db.Column(db.Float, default=0)


class Payroll(db.Model):
    __tablename__ = "payrolls"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id"),
        nullable=False
    )

    period_start = db.Column(db.Date, nullable=False)
    period_end = db.Column(db.Date, nullable=False)
    pay_date = db.Column(db.Date)

    basic_pay = db.Column(db.Float, default=0)
    worked_days = db.Column(db.Integer, default=0)
    regular_hours = db.Column(db.Float, default=0)
    late_hours = db.Column(db.Float, default=0)
    undertime_hours = db.Column(db.Float, default=0)
    undertime_deduction = db.Column(db.Float, default=0)
    late_deduction = db.Column(db.Float, default=0)
    overtime_hours = db.Column(db.Float, default=0)
    overtime_pay = db.Column(db.Float, default=0)
    regular_ot_hours = db.Column(db.Float, default=0)
    regular_ot_pay = db.Column(db.Float, default=0)
    holiday_ot_hours = db.Column(db.Float, default=0)
    holiday_ot_pay = db.Column(db.Float, default=0)
    rest_day_ot_hours = db.Column(db.Float, default=0)
    rest_day_ot_pay = db.Column(db.Float, default=0)
    holiday_pay = db.Column(db.Float, default=0)
    rest_day_premium = db.Column(db.Float, default=0)
    absent_days = db.Column(db.Integer, default=0)
    leave_with_pay_days = db.Column(db.Integer, default=0)
    leave_without_pay_days = db.Column(db.Integer, default=0)
    vl_days = db.Column(db.Integer, default=0)
    sl_days = db.Column(db.Integer, default=0)
    vl_pay = db.Column(db.Float, default=0)
    sl_pay = db.Column(db.Float, default=0)
    vl_conversion = db.Column(db.Float, default=0)
    sl_conversion = db.Column(db.Float, default=0)
    allowance = db.Column(db.Float, default=0)
    incentive = db.Column(db.Float, default=0)
    gross_pay = db.Column(db.Float, default=0)

    sss = db.Column(db.Float, default=0)
    philhealth = db.Column(db.Float, default=0)
    pagibig = db.Column(db.Float, default=0)
    withholding_tax = db.Column(db.Float, default=0)
    loan = db.Column(db.Float, default=0)
    loan_sss_salary = db.Column(db.Float, default=0)
    loan_sss_calamity = db.Column(db.Float, default=0)
    loan_pagibig_mpl = db.Column(db.Float, default=0)
    loan_pagibig_calamity = db.Column(db.Float, default=0)
    cash_advance = db.Column(db.Float, default=0)
    other_deduction = db.Column(db.Float, default=0)
    adjustment_addition = db.Column(db.Float, default=0)
    adjustment_deduction = db.Column(db.Float, default=0)
    adjustment_note = db.Column(db.Text)
    total_deductions = db.Column(db.Float, default=0)

    net_pay = db.Column(db.Float, default=0)
    status = db.Column(db.String(30), default="Draft")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship("Employee", back_populates="payrolls")
