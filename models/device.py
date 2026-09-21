from datetime import datetime
from extensions import db


class Device(db.Model):
    __tablename__ = "devices"

    id = db.Column(db.Integer, primary_key=True)
    device_id = db.Column(db.String(64), unique=True, nullable=False, index=True)
    device_name = db.Column(db.String(120))
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=True)
    status = db.Column(db.String(20), default="Pending", nullable=False)
    user_agent_hash = db.Column(db.String(64))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_seen_at = db.Column(db.DateTime)
    approved_at = db.Column(db.DateTime)

    employee = db.relationship("Employee", backref=db.backref("devices", lazy=True))
