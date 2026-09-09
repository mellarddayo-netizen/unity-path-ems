from datetime import datetime
from extensions import db


class FinalPay(db.Model):
    __tablename__ = "final_pays"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False, unique=True)

    separation_date = db.Column(db.Date, nullable=False)
    separation_type = db.Column(db.String(50), default="Resignation")

    last_payroll_end = db.Column(db.Date)
    coverage_start = db.Column(db.Date)
    coverage_end = db.Column(db.Date)

    basic_pay = db.Column(db.Float, default=0)
    overtime_hours = db.Column(db.Float, default=0)
    overtime_pay = db.Column(db.Float, default=0)
    thirteenth_month = db.Column(db.Float, default=0)
    gross_pay = db.Column(db.Float, default=0)

    sss_employee = db.Column(db.Float, default=0)
    philhealth_employee = db.Column(db.Float, default=0)
    pagibig_employee = db.Column(db.Float, default=0)
    withholding_tax = db.Column(db.Float, default=0)

    requested_statutory = db.Column(db.Float, default=0)
    actual_deductions = db.Column(db.Float, default=0)
    uncollected_deductions = db.Column(db.Float, default=0)
    net_pay = db.Column(db.Float, default=0)

    sss_employer = db.Column(db.Float, default=0)
    philhealth_employer = db.Column(db.Float, default=0)
    pagibig_employer = db.Column(db.Float, default=0)

    status = db.Column(db.String(30), default="Draft")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship("Employee", backref=db.backref("final_pay", uselist=False))
