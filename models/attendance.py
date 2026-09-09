from datetime import datetime, date
from extensions import db


class Attendance(db.Model):
    __tablename__ = "attendances"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id"),
        nullable=False
    )

    # DATE only. Never use datetime.utcnow as the default here.
    attendance_date = db.Column(db.Date, nullable=False, default=date.today)

    time_in = db.Column(db.DateTime)
    time_out = db.Column(db.DateTime)

    total_hours = db.Column(db.Float, default=0)
    late_hours = db.Column(db.Float, default=0)
    undertime_hours = db.Column(db.Float, default=0)
    overtime_hours = db.Column(db.Float, default=0)
    # OT actually approved for payment by Admin. Raw overtime_hours is only the time beyond 5PM.
    approved_overtime_hours = db.Column(db.Float, default=0)

    status = db.Column(db.String(30), default="Present")

    employee = db.relationship("Employee", back_populates="attendances")
