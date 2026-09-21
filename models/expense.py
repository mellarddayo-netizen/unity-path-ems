from datetime import datetime
from extensions import db

class Expense(db.Model):
    __tablename__ = "expenses"

    id = db.Column(db.Integer, primary_key=True)
    expense_date = db.Column(db.Date, nullable=False)
    name = db.Column(db.String(150), nullable=False)
    category = db.Column(db.String(80), nullable=False, default="Other Expense")
    amount = db.Column(db.Float, nullable=False, default=0)
    due_date = db.Column(db.Date, nullable=True)
    paid_date = db.Column(db.Date, nullable=True)
    status = db.Column(db.String(30), nullable=False, default="Pending")
    remarks = db.Column(db.Text, nullable=True)
    source_type = db.Column(db.String(40), nullable=True)
    source_key = db.Column(db.String(100), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    __table_args__ = (
        db.UniqueConstraint("source_type", "source_key", name="uq_expense_source"),
    )
