from datetime import datetime
from extensions import db


class FinancialIncome(db.Model):
    __tablename__ = "financial_income"

    id = db.Column(db.Integer, primary_key=True)
    income_date = db.Column(db.Date, nullable=False)
    name = db.Column(db.String(150), nullable=False)
    category = db.Column(db.String(80), nullable=False, default="Other Income")
    amount = db.Column(db.Float, nullable=False, default=0)
    remarks = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
