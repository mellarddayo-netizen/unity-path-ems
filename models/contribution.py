from datetime import datetime
from extensions import db


class MonthlyContribution(db.Model):
    __tablename__ = "monthly_contributions"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False)
    contribution_year = db.Column(db.Integer, nullable=False)
    contribution_month = db.Column(db.Integer, nullable=False)

    sss_remuneration = db.Column(db.Float, default=0)
    sss_msc = db.Column(db.Float, default=0)
    philhealth_basic_salary = db.Column(db.Float, default=0)
    pagibig_monthly_compensation = db.Column(db.Float, default=0)

    sss_due = db.Column(db.Float, default=0)
    philhealth_due = db.Column(db.Float, default=0)
    pagibig_due = db.Column(db.Float, default=0)

    sss_collected = db.Column(db.Float, default=0)
    philhealth_collected = db.Column(db.Float, default=0)
    pagibig_collected = db.Column(db.Float, default=0)

    sss_uncollected = db.Column(db.Float, default=0)
    philhealth_uncollected = db.Column(db.Float, default=0)
    pagibig_uncollected = db.Column(db.Float, default=0)

    sss_employer = db.Column(db.Float, default=0)
    philhealth_employer = db.Column(db.Float, default=0)
    pagibig_employer = db.Column(db.Float, default=0)

    status = db.Column(db.String(30), default="Open")
    source_payroll_id = db.Column(db.Integer, db.ForeignKey("payrolls.id"))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    employee = db.relationship("Employee", backref=db.backref("monthly_contributions", lazy=True))
    source_payroll = db.relationship("Payroll", backref=db.backref("monthly_contribution", uselist=False))

    __table_args__ = (
        db.UniqueConstraint(
            "employee_id", "contribution_year", "contribution_month",
            name="uq_monthly_contribution_employee_period"
        ),
    )
