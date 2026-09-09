from datetime import datetime
from extensions import db
from werkzeug.security import generate_password_hash, check_password_hash


class Employee(db.Model):
    __tablename__ = "employees"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(50), unique=True, nullable=False)
    first_name = db.Column(db.String(100), nullable=False)
    middle_name = db.Column(db.String(100))
    last_name = db.Column(db.String(100), nullable=False)

    email = db.Column(db.String(150))
    phone = db.Column(db.String(50))
    address = db.Column(db.Text)
    department = db.Column(db.String(100))
    position = db.Column(db.String(100))
    date_hired = db.Column(db.Date)
    regularization_date = db.Column(db.Date)
    vl_entitlement = db.Column(db.Float, default=0)
    sl_entitlement = db.Column(db.Float, default=0)
    employment_status = db.Column(db.String(30), default="Active")
    daily_rate = db.Column('salary', db.Float, default=0)
    daily_allowance = db.Column(db.Float, default=0)
    daily_incentive = db.Column(db.Float, default=0)
    loan_deduction = db.Column(db.Float, default=0)
    separation_date = db.Column(db.Date)
    separation_type = db.Column(db.String(50))
    separation_reason = db.Column(db.Text)

    profile_photo = db.Column(db.String(255))

    sss_number = db.Column(db.String(50))
    philhealth_number = db.Column(db.String(50))
    pagibig_number = db.Column(db.String(50))
    tin_number = db.Column(db.String(50))
    umid_number = db.Column(db.String(50))
    national_id = db.Column(db.String(80))
    voters_id = db.Column(db.String(80))
    drivers_license = db.Column(db.String(80))
    passport_number = db.Column(db.String(80))
    passport_expiry = db.Column(db.String(30))

    bank_name = db.Column(db.String(100))
    bank_account_name = db.Column(db.String(150))
    bank_account_number = db.Column(db.String(100))

    emergency_name = db.Column(db.String(150))
    emergency_relationship = db.Column(db.String(100))
    emergency_phone = db.Column(db.String(50))
    emergency_address = db.Column(db.Text)

    # Extended HR / 201-file profile information
    nickname = db.Column(db.String(100))
    date_of_birth = db.Column(db.String(30))
    place_of_birth = db.Column(db.String(200))
    sex = db.Column(db.String(30))
    civil_status = db.Column(db.String(50))
    nationality = db.Column(db.String(100))
    religion = db.Column(db.String(100))
    blood_type = db.Column(db.String(10))
    height = db.Column(db.String(30))
    weight = db.Column(db.String(30))
    permanent_address = db.Column(db.Text)
    secondary_phone = db.Column(db.String(50))
    company_email = db.Column(db.String(150))
    social_media = db.Column(db.String(255))
    professional_summary = db.Column(db.Text)
    reporting_to = db.Column(db.String(150))
    work_location = db.Column(db.String(255))
    employment_type = db.Column(db.String(50), default="Full-Time")
    schedule_type = db.Column(db.String(50), default="Fixed")
    schedule_details = db.Column(db.String(255))
    education_json = db.Column(db.Text)
    work_experience_json = db.Column(db.Text)
    skills_json = db.Column(db.Text)
    family_json = db.Column(db.Text)
    references_json = db.Column(db.Text)
    documents_json = db.Column(db.Text)
    duties = db.Column(db.Text)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    leave_requests = db.relationship(
        "LeaveRequest",
        back_populates="employee",
        cascade="all, delete-orphan",
        lazy=True
    )

    attendances = db.relationship(
        "Attendance",
        back_populates="employee",
        cascade="all, delete-orphan",
        lazy=True
    )

    payrolls = db.relationship(
        "Payroll",
        back_populates="employee",
        cascade="all, delete-orphan",
        lazy=True
    )

    user_account = db.relationship(
        "User",
        back_populates="employee",
        uselist=False,
        cascade="all, delete-orphan"
    )


class User(db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(30), default="employee")

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id"),
        nullable=True
    )

    employee = db.relationship("Employee", back_populates="user_account")

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
