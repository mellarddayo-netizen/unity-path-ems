from datetime import datetime
from extensions import db


class CollectionClient(db.Model):
    __tablename__ = "collection_clients"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    contact_person = db.Column(db.String(150))
    email = db.Column(db.String(150))
    phone = db.Column(db.String(50))
    commission_rate = db.Column(db.Float, default=0)
    status = db.Column(db.String(30), default="Active")
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    accounts = db.relationship("CollectionAccount", back_populates="client", cascade="all, delete-orphan")


class CollectionDebtor(db.Model):
    __tablename__ = "collection_debtors"
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(200), nullable=False)
    account_number = db.Column(db.String(100))
    phone = db.Column(db.String(50))
    secondary_phone = db.Column(db.String(50))
    email = db.Column(db.String(150))
    address = db.Column(db.Text)
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    accounts = db.relationship("CollectionAccount", back_populates="debtor")


class CollectionAccount(db.Model):
    __tablename__ = "collection_accounts"
    id = db.Column(db.Integer, primary_key=True)
    account_number = db.Column(db.String(100), unique=True, nullable=False)
    client_id = db.Column(db.Integer, db.ForeignKey("collection_clients.id"), nullable=False)
    debtor_id = db.Column(db.Integer, db.ForeignKey("collection_debtors.id"), nullable=False)
    assigned_employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"))
    principal_balance = db.Column(db.Float, default=0)
    interest_balance = db.Column(db.Float, default=0)
    penalty_balance = db.Column(db.Float, default=0)
    outstanding_balance = db.Column(db.Float, default=0)
    due_date = db.Column(db.Date)
    status = db.Column(db.String(40), default="Assigned")
    priority = db.Column(db.String(20), default="Normal")
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    client = db.relationship("CollectionClient", back_populates="accounts")
    debtor = db.relationship("CollectionDebtor", back_populates="accounts")
    assigned_employee = db.relationship("Employee", foreign_keys=[assigned_employee_id])
    activities = db.relationship("CollectionActivity", back_populates="account", cascade="all, delete-orphan")
    ptps = db.relationship("PromiseToPay", back_populates="account", cascade="all, delete-orphan")
    payments = db.relationship("CollectionPayment", back_populates="account", cascade="all, delete-orphan")


class CollectionActivity(db.Model):
    __tablename__ = "collection_activities"
    id = db.Column(db.Integer, primary_key=True)
    account_id = db.Column(db.Integer, db.ForeignKey("collection_accounts.id"), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"))
    activity_type = db.Column(db.String(50), nullable=False)
    outcome = db.Column(db.String(100))
    remarks = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    account = db.relationship("CollectionAccount", back_populates="activities")
    employee = db.relationship("Employee")


class PromiseToPay(db.Model):
    __tablename__ = "collection_ptps"
    id = db.Column(db.Integer, primary_key=True)
    account_id = db.Column(db.Integer, db.ForeignKey("collection_accounts.id"), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"))
    promise_date = db.Column(db.Date, nullable=False)
    promise_amount = db.Column(db.Float, default=0)
    status = db.Column(db.String(30), default="Pending")
    remarks = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    account = db.relationship("CollectionAccount", back_populates="ptps")
    employee = db.relationship("Employee")


class CollectionPayment(db.Model):
    __tablename__ = "collection_payments"
    id = db.Column(db.Integer, primary_key=True)
    account_id = db.Column(db.Integer, db.ForeignKey("collection_accounts.id"), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"))
    amount = db.Column(db.Float, default=0)
    payment_date = db.Column(db.Date, nullable=False)
    payment_method = db.Column(db.String(50))
    reference_no = db.Column(db.String(100))
    remarks = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    account = db.relationship("CollectionAccount", back_populates="payments")
    employee = db.relationship("Employee")


class CollectionCommission(db.Model):
    __tablename__ = "collection_commissions"
    id = db.Column(db.Integer, primary_key=True)
    payment_id = db.Column(db.Integer, db.ForeignKey("collection_payments.id"), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False)
    rate = db.Column(db.Float, default=0)
    amount = db.Column(db.Float, default=0)
    status = db.Column(db.String(30), default="Pending")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    payment = db.relationship("CollectionPayment")
    employee = db.relationship("Employee")
