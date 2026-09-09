from datetime import datetime
from extensions import db


class LeaveRequest(db.Model):
    __tablename__ = "leave_requests"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False)
    leave_type = db.Column(db.String(10), nullable=False)  # VL / SL
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    reason = db.Column(db.Text)
    status = db.Column(db.String(20), default="Pending", nullable=False)
    admin_remarks = db.Column(db.Text)
    approved_by = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    approved_at = db.Column(db.DateTime)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship("Employee", back_populates="leave_requests")
    approver = db.relationship("User", foreign_keys=[approved_by])
