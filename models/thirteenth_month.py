from datetime import datetime
from extensions import db


class ThirteenthMonthRecord(db.Model):
    __tablename__ = "thirteenth_month_records"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False)
    year = db.Column(db.Integer, nullable=False)
    basic_salary_earned = db.Column(db.Float, default=0)
    thirteenth_amount = db.Column(db.Float, default=0)
    status = db.Column(db.String(30), default="Pending", nullable=False)
    due_date = db.Column(db.Date, nullable=True)
    paid_date = db.Column(db.Date, nullable=True)
    payment_reference = db.Column(db.String(150), nullable=True)
    remarks = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    employee = db.relationship("Employee", backref=db.backref("thirteenth_month_records", lazy=True))

    __table_args__ = (
        db.UniqueConstraint("employee_id", "year", name="uq_13th_month_employee_year"),
    )
