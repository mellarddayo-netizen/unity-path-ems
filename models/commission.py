from datetime import datetime
from extensions import db


class CommissionRecord(db.Model):
    __tablename__ = "commission_records"

    id = db.Column(db.Integer, primary_key=True)
    period = db.Column(db.String(50), nullable=False, index=True)
    agent_name = db.Column(db.String(200), nullable=False)
    employee_id = db.Column(db.String(50), nullable=False, index=True)
    total_collection = db.Column(db.Float, default=0, nullable=False)
    commission_rate = db.Column(db.Float, default=0, nullable=False)
    commission = db.Column(db.Float, default=0, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
