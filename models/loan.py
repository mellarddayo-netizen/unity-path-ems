from datetime import datetime
from extensions import db


class EmployeeLoan(db.Model):
    __tablename__ = "employee_loans"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id"),
        nullable=False
    )
    loan_type = db.Column(db.String(100), nullable=False)
    reference_no = db.Column(db.String(100))
    original_amount = db.Column(db.Float, default=0)
    balance = db.Column(db.Float, default=0)
    deduction_per_payroll = db.Column(db.Float, default=0)
    start_date = db.Column(db.Date)
    end_date = db.Column(db.Date)
    company_remittance = db.Column(db.Boolean, default=True)
    status = db.Column(db.String(30), default="Active")
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship("Employee", backref=db.backref("loans", lazy=True))
    payments = db.relationship(
        "LoanPayment",
        back_populates="loan",
        cascade="all, delete-orphan",
        lazy=True
    )


class LoanPayment(db.Model):
    __tablename__ = "loan_payments"

    id = db.Column(db.Integer, primary_key=True)
    loan_id = db.Column(
        db.Integer,
        db.ForeignKey("employee_loans.id"),
        nullable=False
    )
    payroll_id = db.Column(
        db.Integer,
        db.ForeignKey("payrolls.id"),
        nullable=False
    )
    amount = db.Column(db.Float, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    loan = db.relationship("EmployeeLoan", back_populates="payments")
    payroll = db.relationship("Payroll", backref=db.backref("loan_payments", lazy=True))
