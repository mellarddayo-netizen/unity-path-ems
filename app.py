import calendar
import math
import os
import json
import smtplib
import uuid
import hashlib
from email.message import EmailMessage
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime, date, time, timedelta
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, abort, send_file, send_from_directory, g
)
from werkzeug.utils import secure_filename
from sqlalchemy import or_, inspect
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from dotenv import load_dotenv

from extensions import db

from models import Device, Employee, User, Attendance, Payroll, PayrollSettings, FinalPay, EmployeeLoan, LoanPayment, MonthlyContribution, Holiday, LeaveRequest, CommissionRecord, Expense, FinancialIncome, ThirteenthMonthRecord


# Standard company attendance schedule
REGULAR_TIME_IN = time(8, 0)
REGULAR_TIME_OUT = time(17, 0)
UNPAID_LUNCH_START = time(12, 0)
UNPAID_LUNCH_END = time(13, 0)
MAX_REGULAR_PAID_HOURS = 8.0

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))
DB_PATH = os.path.join(BASE_DIR, "employee_system_v10.db")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = "postgresql://" + DATABASE_URL[len("postgres://"):]
PROFILE_STORAGE_DIR = os.getenv("PROFILE_STORAGE_DIR", "").strip()
PROFILE_DIR = PROFILE_STORAGE_DIR or os.path.join(BASE_DIR, "static", "uploads", "profiles")

ALLOWED_IMAGE_EXTENSIONS = {"jpg", "jpeg", "png", "webp"}

WORK_START = time(8, 0)
WORK_END = time(17, 0)
BREAK_HOURS = 1.0
REQUIRED_WORK_HOURS = 8.0


app = Flask(__name__)
app.config["SECRET_KEY"] = "change-this-secret-key"
app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL or ("sqlite:///" + DB_PATH)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

db.init_app(app)
app.jinja_env.globals["timedelta"] = timedelta

os.makedirs(PROFILE_DIR, exist_ok=True)


DEVICE_COOKIE_NAME = "unity_path_device"
DEVICE_COOKIE_MAX_AGE = 60 * 60 * 24 * 365 * 2  # 2 years


def _user_agent_hash():
    return hashlib.sha256((request.headers.get("User-Agent", "") or "").encode("utf-8")).hexdigest()


def get_device_id():
    """Return the browser/device registration ID for this browser.

    This is intentionally a web-device ID, not a MAC address. Browsers do not
    expose the laptop's MAC address to normal web applications.
    """
    device_id = request.cookies.get(DEVICE_COOKIE_NAME)
    if device_id:
        return device_id
    device_id = uuid.uuid4().hex.upper()
    g.new_device_id = device_id
    return device_id


@app.after_request
def persist_device_cookie(response):
    device_id = getattr(g, "new_device_id", None)
    if device_id:
        response.set_cookie(
            DEVICE_COOKIE_NAME,
            device_id,
            max_age=DEVICE_COOKIE_MAX_AGE,
            httponly=True,
            secure=not app.debug,
            samesite="Lax",
        )
    return response


def current_device():
    device_id = get_device_id()
    return Device.query.filter_by(device_id=device_id).first()


def device_is_allowed(user=None):
    """Allow access when this browser/device is on the company whitelist.

    Device approval is intentionally NOT tied to an employee account. Any
    valid EMS employee/admin account may use any Active company device.
    """
    device = current_device()
    if not device or device.status != "Active":
        return False
    if device.user_agent_hash and device.user_agent_hash != _user_agent_hash():
        return False
    device.last_seen_at = datetime.utcnow()
    return True


def register_or_update_device(device_id, user=None, device_name=None, status="Active"):
    """Create/update a company-whitelisted device.

    The device is company-owned/authorized rather than employee-owned, so
    employee_id is deliberately not assigned here.
    """
    device = Device.query.filter_by(device_id=device_id).first()
    if not device:
        device = Device(
            device_id=device_id,
            device_name=(device_name or "Company Device")[:120],
            employee_id=None,
            status=status,
            user_agent_hash=_user_agent_hash(),
            approved_at=datetime.utcnow() if status == "Active" else None,
            last_seen_at=datetime.utcnow(),
        )
        db.session.add(device)
    else:
        if device_name:
            device.device_name = device_name[:120]
        device.status = status
        device.user_agent_hash = _user_agent_hash()
        device.approved_at = datetime.utcnow() if status == "Active" else device.approved_at
        device.last_seen_at = datetime.utcnow()
        device.employee_id = None
    return device




def create_pending_device_for_user(user=None):
    """Create an unassigned pending company device for admin approval."""
    device_id = get_device_id()
    device = Device.query.filter_by(device_id=device_id).first()
    if not device:
        device = Device(
            device_id=device_id,
            device_name="Pending Company Device",
            employee_id=None,
            status="Pending",
            user_agent_hash=_user_agent_hash(),
            last_seen_at=datetime.utcnow(),
        )
        db.session.add(device)
        db.session.commit()
    return device



@app.route("/profile-photos/<path:filename>", endpoint="profile_photo")
def profile_photo(filename):
    """Serve employee profile photos from the persistent profile storage directory."""
    safe_name = secure_filename(os.path.basename(filename))
    if not safe_name:
        abort(404)
    return send_from_directory(PROFILE_DIR, safe_name)


def admin_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if session.get("role") != "admin":
            flash("Admin access required.", "danger")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapper


def employee_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if session.get("role") != "employee":
            flash("Employee access required.", "danger")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapper


def parse_date(value):
    if not value:
        return None
    return datetime.strptime(value, "%Y-%m-%d").date()


def parse_datetime_local(value):
    if not value:
        return None
    return datetime.strptime(value, "%Y-%m-%dT%H:%M")


def parse_time(value):
    if not value:
        return None
    return datetime.strptime(value, "%H:%M").time()


def hours_between(start, end):
    if not start or not end:
        return 0.0
    seconds = (end - start).total_seconds()
    return max(0.0, seconds / 3600)


def get_holiday(holiday_date):
    return Holiday.query.filter_by(holiday_date=holiday_date, active=True).first()


def holiday_defaults(holiday_type):
    """Default Philippine payroll multipliers; editable per holiday in Calendar."""
    mapping = {
        "Regular Holiday": (2.00, 1.00, 1.30),
        "Special Non-Working Day": (1.30, 0.00, 1.30),
        "Special Working Day": (1.00, 0.00, 1.30),
        "Company Holiday": (1.00, 1.00, 1.30),
        "Rest Day": (1.30, 0.00, 1.30),
    }
    return mapping.get(holiday_type, (1.00, 0.00, 1.30))


def calculate_holiday_earnings(employee, period_start, period_end, settings):
    """Calculate holiday and Sunday-rest premiums from the Calendar.

    Company schedule:
      - Monday-Friday: regular workdays at 100% when worked
      - Saturday: optional workday; when worked, 100% regular rate
      - Sunday: the only official scheduled rest day

    The Calendar is the source of truth for holiday classification and
    configured multipliers. Holiday premium and Sunday rest-day premium are
    returned separately so the payslip cannot label a Sunday premium as a
    holiday when there is no holiday on that date.
    """
    records = payable_attendance_records(employee.id, period_start, period_end)
    by_date = {r.attendance_date: r for r in records}
    daily = max(0.0, float(employee.daily_rate or 0))
    hourly = daily / max(float(settings.hours_per_day or 8), 1.0)

    holiday_premium = 0.0
    rest_day_premium = 0.0
    holiday_ot_pay = 0.0
    rest_day_ot_pay = 0.0
    holiday_days = []

    d = period_start
    while d <= period_end:
        h = get_holiday(d)
        rec = by_date.get(d)
        is_sunday = d.weekday() == 6

        if h:
            work_mult = max(0.0, float(h.work_multiplier or 0))
            unworked_mult = max(0.0, float(h.unworked_multiplier or 0))
            ot_mult = max(0.0, float(h.ot_multiplier or 1.30))

            if is_sunday:
                if h.holiday_type == "Regular Holiday":
                    work_mult = 2.60
                    holiday_component_mult = 1.00
                    rest_component_mult = 0.60
                elif h.holiday_type == "Special Non-Working Day":
                    work_mult = 1.50
                    holiday_component_mult = 0.30
                    rest_component_mult = 0.20
                else:
                    work_mult = max(work_mult, 1.30)
                    holiday_component_mult = max(0.0, min(work_mult, 1.30) - 1.0)
                    rest_component_mult = max(0.0, work_mult - 1.30)
            else:
                holiday_component_mult = max(0.0, work_mult - 1.0)
                rest_component_mult = 0.0

            if rec:
                if holiday_component_mult:
                    holiday_premium += daily * holiday_component_mult
                if rest_component_mult:
                    rest_day_premium += daily * rest_component_mult

                approved_ot = max(0.0, float(rec.approved_overtime_hours or 0))
                if approved_ot:
                    ot_pay = approved_ot * hourly * work_mult * ot_mult
                    if is_sunday:
                        rest_day_ot_pay += ot_pay
                    else:
                        holiday_ot_pay += ot_pay
                holiday_days.append((d, h.name, h.holiday_type, work_mult, approved_ot))
            elif unworked_mult > 0:
                # Unworked holiday pay remains a separate holiday earning.
                holiday_premium += daily * unworked_mult
                holiday_days.append((d, h.name, h.holiday_type, unworked_mult, 0.0))

        elif rec and is_sunday:
            work_mult = 1.30
            rest_day_premium += daily * (work_mult - 1.0)
            approved_ot = max(0.0, float(rec.approved_overtime_hours or 0))
            if approved_ot:
                rest_day_ot_pay += approved_ot * hourly * work_mult * 1.30
            holiday_days.append((d, "Sunday Rest Day", "Rest Day", work_mult, approved_ot))

        d += timedelta(days=1)

    ordinary_ot_hours = 0.0
    for rec in records:
        if not get_holiday(rec.attendance_date) and rec.attendance_date.weekday() != 6:
            ordinary_ot_hours += max(0.0, float(rec.approved_overtime_hours or 0))

    ordinary_ot_pay = calculate_overtime_pay(daily, ordinary_ot_hours, settings)
    holiday_ot_hours = round(sum(x[4] for x in holiday_days if x[0].weekday() != 6), 2)
    rest_day_ot_hours = round(sum(x[4] for x in holiday_days if x[0].weekday() == 6), 2)
    return {
        "holiday_premium": round(holiday_premium, 2),
        "rest_day_premium": round(rest_day_premium, 2),
        "holiday_adjustment": round(holiday_premium + rest_day_premium, 2),
        "holiday_ot_pay": round(holiday_ot_pay, 2),
        "rest_day_ot_pay": round(rest_day_ot_pay, 2),
        "ordinary_ot_hours": round(ordinary_ot_hours, 2),
        "ordinary_ot_pay": round(ordinary_ot_pay, 2),
        "holiday_ot_hours": holiday_ot_hours,
        "rest_day_ot_hours": rest_day_ot_hours,
        "holiday_days": holiday_days,
    }


def get_attendance_summary(employee_id, period_start, period_end):
    """Return transparent attendance counts for a payroll period.

    Absence is counted only on scheduled Monday-Friday dates with no valid
    attendance record and no active Calendar holiday. Saturday is optional
    work and Sunday is the official rest day, so neither is an absence.
    """
    valid = {r.attendance_date: r for r in payable_attendance_records(employee_id, period_start, period_end)}
    absent_days = 0
    leave_with_pay_days = 0
    leave_without_pay_days = 0
    d = period_start
    while d <= period_end:
        h = get_holiday(d)
        rec = valid.get(d)
        if d.weekday() < 5 and not h:
            if rec is None:
                absent_days += 1
        if rec:
            status = (rec.status or "").strip().lower()
            if status == "leave with pay":
                leave_with_pay_days += 1
            elif status == "leave without pay":
                leave_without_pay_days += 1
        d += timedelta(days=1)
    return {
        "absent_days": absent_days,
        "leave_with_pay_days": leave_with_pay_days,
        "leave_without_pay_days": leave_without_pay_days,
    }

def calculate_attendance(record):
    """
    Attendance rules:
      - Regular schedule: 8:00 AM to 5:00 PM.
      - 12:00 PM to 1:00 PM is unpaid lunch.
      - Paid regular hours are capped at 8 hours/day.
      - Late is measured strictly from 8:00 AM.
      - Undertime is measured strictly against the 5:00 PM scheduled end.
        Early arrival does NOT offset undertime.
      - Overtime is time worked after 5:00 PM.
      - Time before 8:00 AM does NOT count as regular paid hours.
    """
    record.total_hours = 0.0
    record.late_hours = 0.0
    record.undertime_hours = 0.0
    record.overtime_hours = 0.0

    # VL/SL are paid leave records entered by Admin. They do not require Time In/Out.
    if (record.status or "").upper() in ("VL", "SL") and not record.time_in and not record.time_out:
        record.total_hours = 0.0
        record.late_hours = 0.0
        record.undertime_hours = 0.0
        record.overtime_hours = 0.0
        record.approved_overtime_hours = 0.0
        record.status = (record.status or "").upper()
        return

    # A Time Out is valid only when it belongs to the same calendar date
    # as the attendance record / Time In. A cross-day or orphan Time Out
    # must never be used for payroll calculations.
    if record.time_out and (
        not record.time_in
        or record.time_out.date() != record.attendance_date
        or record.time_out.date() != record.time_in.date()
    ):
        record.approved_overtime_hours = 0.0
        record.status = "Invalid Out"
        return

    if record.time_in:
        scheduled_start = datetime.combine(record.attendance_date, WORK_START)
        if record.time_in > scheduled_start:
            record.late_hours = round(
                (record.time_in - scheduled_start).total_seconds() / 3600.0, 2
            )

    if record.time_in and record.time_out:
        regular_start = datetime.combine(record.attendance_date, WORK_START)
        regular_end = datetime.combine(record.attendance_date, WORK_END)

        # Paid regular time is only within the 8AM-5PM window.
        effective_start = max(record.time_in, regular_start)
        effective_end = min(record.time_out, regular_end)

        regular_elapsed = 0.0
        if effective_end > effective_start:
            regular_elapsed = (
                effective_end - effective_start
            ).total_seconds() / 3600.0

        # Subtract only the actual overlap with the unpaid 12PM-1PM lunch.
        lunch_start = datetime.combine(record.attendance_date, time(12, 0))
        lunch_end = datetime.combine(record.attendance_date, time(13, 0))

        lunch_overlap = 0.0
        overlap_start = max(effective_start, lunch_start)
        overlap_end = min(effective_end, lunch_end)
        if overlap_end > overlap_start:
            lunch_overlap = (
                overlap_end - overlap_start
            ).total_seconds() / 3600.0

        paid_regular = max(0.0, regular_elapsed - lunch_overlap)

        # Maximum paid regular time is 8 hours.
        record.total_hours = round(
            min(REQUIRED_WORK_HOURS, paid_regular), 2
        )

        # Strict undertime: based on the scheduled 5PM end, not on
        # whether the employee arrived early enough to "make up" the hours.
        if record.time_out < regular_end:
            record.undertime_hours = round(
                (regular_end - record.time_out).total_seconds() / 3600.0,
                2
            )

        # OT starts strictly after 5PM.
        if record.time_out > regular_end:
            record.overtime_hours = round(
                (record.time_out - regular_end).total_seconds() / 3600.0,
                2
            )

        # Keep previous approval only up to the new actual OT; never auto-approve.
        record.approved_overtime_hours = round(
            min(
                max(0.0, float(record.approved_overtime_hours or 0)),
                max(0.0, float(record.overtime_hours or 0))
            ),
            2
        )

        record.status = "Present"
    elif record.time_in:
        record.status = "Time In"
    else:
        record.status = "Absent"


def is_valid_payable_attendance(record):
    """True only for a completed same-day Present attendance record."""
    if not record or not record.time_in or not record.time_out:
        return False
    if record.status != "Present":
        return False
    if record.time_in.date() != record.attendance_date:
        return False
    if record.time_out.date() != record.attendance_date:
        return False
    return record.time_out >= record.time_in


def payable_attendance_records(employee_id, period_start, period_end):
    """Return one valid payable attendance record per calendar date."""
    records = Attendance.query.filter(
        Attendance.employee_id == employee_id,
        Attendance.attendance_date >= period_start,
        Attendance.attendance_date <= period_end
    ).all()
    by_date = {}
    for record in records:
        if not is_valid_payable_attendance(record):
            continue
        current = by_date.get(record.attendance_date)
        if current is None or float(record.total_hours or 0) > float(current.total_hours or 0):
            by_date[record.attendance_date] = record
    return list(by_date.values())


def get_overtime_hours(employee_id, start, end):
    """Return only Admin-approved OT from valid Present attendance."""
    records = payable_attendance_records(employee_id, start, end)
    return round(sum(float(r.approved_overtime_hours or 0) for r in records), 2)

def month_periods(year, month):
    last_day = calendar.monthrange(year, month)[1]
    return (
        date(year, month, 1),
        date(year, month, 15),
        date(year, month, 16),
        date(year, month, last_day),
    )


def get_attendance_hours(employee_id, period_start, period_end):
    records = payable_attendance_records(employee_id, period_start, period_end)
    return round(sum(max(0.0, float(r.total_hours or 0.0)) for r in records), 2)


def get_worked_days(employee_id, period_start, period_end):
    """Count only completed, same-day Present attendance days."""
    return len(payable_attendance_records(employee_id, period_start, period_end))

def get_leave_usage(employee, year):
    """Return annual VL/SL usage and remaining balances.

    VL/SL are recorded through Attendance using status "VL" or "SL".
    Entitlements are entered by Admin in Employee Details upon regularization.
    Usage is counted from the employee regularization date (or Jan 1 if none).
    """
    start = date(year, 1, 1)
    if getattr(employee, "regularization_date", None) and employee.regularization_date > start:
        start = employee.regularization_date
    end = date(year, 12, 31)
    records = Attendance.query.filter(
        Attendance.employee_id == employee.id,
        Attendance.attendance_date >= start,
        Attendance.attendance_date <= end,
        Attendance.status.in_(["VL", "SL"])
    ).all()
    vl_used = sum(1 for r in records if (r.status or "").upper() == "VL")
    sl_used = sum(1 for r in records if (r.status or "").upper() == "SL")
    vl_ent = max(0.0, float(getattr(employee, "vl_entitlement", 0) or 0))
    sl_ent = max(0.0, float(getattr(employee, "sl_entitlement", 0) or 0))
    return {
        "vl_entitlement": vl_ent,
        "sl_entitlement": sl_ent,
        "vl_used": vl_used,
        "sl_used": sl_used,
        "vl_balance": max(0.0, vl_ent - vl_used),
        "sl_balance": max(0.0, sl_ent - sl_used),
    }

def leave_days_for_period(employee, period_start, period_end):
    records = Attendance.query.filter(
        Attendance.employee_id == employee.id,
        Attendance.attendance_date >= period_start,
        Attendance.attendance_date <= period_end,
        Attendance.status.in_(["VL", "SL"])
    ).all()
    return {
        "vl_days": sum(1 for r in records if (r.status or "").upper() == "VL"),
        "sl_days": sum(1 for r in records if (r.status or "").upper() == "SL"),
    }

def leave_conversion_for_period(employee, period_start, period_end):
    """Convert all unused VL/SL to cash in the December 2nd cutoff.

    The conversion is calculated from the saved annual entitlement less VL/SL
    actually used during the year, multiplied by the employee's Basic Daily Rate.
    No conversion is made before the December 16-last-day payroll.
    """
    if period_start.month != 12 or period_start.day != 16:
        return {"vl_conversion": 0.0, "sl_conversion": 0.0}
    usage = get_leave_usage(employee, period_start.year)
    daily_rate = max(0.0, float(employee.daily_rate or 0))
    return {
        "vl_conversion": round(usage["vl_balance"] * daily_rate, 2),
        "sl_conversion": round(usage["sl_balance"] * daily_rate, 2),
    }

def calculate_basic_pay(employee, period_start, period_end, settings):
    """
    Scheduled basic pay before late/undertime deductions.

    Each completed attendance day earns the employee's Basic Daily Rate.
    Late and undertime are deducted separately and shown transparently in
    payroll. Early arrival does not offset late/undertime.

    Regular schedule:
      8:00 AM - 5:00 PM
      12:00 PM - 1:00 PM unpaid lunch
      8 paid regular hours/day maximum
    """
    worked_days = get_worked_days(employee.id, period_start, period_end)
    leave_days = leave_days_for_period(employee, period_start, period_end)
    paid_days = worked_days + leave_days["vl_days"] + leave_days["sl_days"]
    return round(
        max(0.0, float(employee.daily_rate or 0)) * paid_days,
        2
    )


def calculate_overtime_pay(daily_rate, overtime_hours, settings):
    """Calculate approved OT pay from the employee daily rate."""
    hourly_rate = float(daily_rate or 0) / max(float(settings.hours_per_day or 8), 1)
    multiplier = max(float(settings.overtime_multiplier or 1.25), 0)
    return round(max(0.0, float(overtime_hours or 0)) * hourly_rate * multiplier, 2)



def _attendance_records_for_period(employee_id, period_start, period_end):
    return payable_attendance_records(employee_id, period_start, period_end)

def _monthly_earnings_basis(employee, year, month):
    """Build statutory bases from the employee's September-style monthly data.

    SSS/Pag-IBIG use actual remuneration/compensation earned in the month.
    PhilHealth uses the employee's fixed monthly basic salary and excludes
    overtime, allowances, absences, tardiness and undertime from its MBS.
    """
    first, _, _, last = month_periods(year, month)
    settings = PayrollSettings.query.first() or PayrollSettings()
    records = _attendance_records_for_period(employee.id, first, last)
    worked_days = len(records)

    daily_rate = max(0.0, float(employee.daily_rate or 0))
    scheduled_basic = round(daily_rate * worked_days, 2)

    late_deduction = 0.0
    undertime_deduction = 0.0
    hourly_rate = daily_rate / max(float(settings.hours_per_day or 8), 1.0)
    late_deduction = round(
        sum(float(r.late_hours or 0) for r in records) * hourly_rate, 2
    )
    undertime_deduction = round(
        sum(float(r.undertime_hours or 0) for r in records) * hourly_rate, 2
    )
    actual_basic_paid = round(
        max(0.0, scheduled_basic - late_deduction - undertime_deduction), 2
    )

    holiday_values = calculate_holiday_earnings(employee, first, last, settings)
    approved_ot_hours = round(
        holiday_values["ordinary_ot_hours"] + sum(x[4] for x in holiday_values["holiday_days"]), 2
    )
    overtime_pay = round(
        holiday_values["ordinary_ot_pay"] + holiday_values["holiday_ot_pay"] + holiday_values["rest_day_ot_pay"], 2
    )
    holiday_pay = round(holiday_values["holiday_premium"], 2)
    rest_day_premium = round(holiday_values["rest_day_premium"], 2)
    allowance = round(max(0.0, float(employee.daily_allowance or 0)) * worked_days, 2)
    incentive = round(max(0.0, float(employee.daily_incentive or 0)) * worked_days, 2)

    adjustment_addition = round(sum(
        float(p.adjustment_addition or 0)
        for p in Payroll.query.filter(
            Payroll.employee_id == employee.id,
            Payroll.period_start >= first,
            Payroll.period_end <= last
        ).all()
    ), 2)

    # SSS actual remuneration includes salaries/wages, overtime and monthly
    # allowances/remuneration. Use the actual payable basic after unpaid
    # tardiness/undertime, not the scheduled basic before those reductions.
    actual_remuneration = round(
        actual_basic_paid + holiday_pay + rest_day_premium + overtime_pay + allowance + incentive + adjustment_addition,
        2
    )

    # PhilHealth MBS is the fixed basic rate. For a daily-rate employee, EMS
    # converts the configured daily rate using the configured monthly working
    # days. Attendance deductions, OT and allowances are excluded from MBS.
    philhealth_basic_salary = round(
        daily_rate * max(float(settings.working_days_per_month or 26), 0), 2
    )

    # Pag-IBIG monthly compensation includes basic salary and other allowances;
    # actual monthly remuneration is used, then capped at P5,000 for the rate.
    pagibig_monthly_compensation = actual_remuneration

    attendance_summary = get_attendance_summary(employee.id, first, last)

    return {
        "worked_days": worked_days,
        "sss_remuneration": actual_remuneration,
        "philhealth_basic_salary": philhealth_basic_salary,
        "pagibig_monthly_compensation": pagibig_monthly_compensation,
        "rest_day_premium": rest_day_premium,
        "holiday_premium": holiday_pay,
        "absent_days": attendance_summary["absent_days"],
        "leave_with_pay_days": attendance_summary["leave_with_pay_days"],
        "leave_without_pay_days": attendance_summary["leave_without_pay_days"],
    }


def _sss_msc(actual_remuneration):
    """Map actual monthly remuneration to the Jan-2025 SSS MSC schedule.

    For employed members, SSS states that MSC is based on total actual
    remuneration from employment. The 2025 schedule uses a P5,000 minimum,
    P35,000 maximum and P500 increments.
    """
    compensation = max(0.0, float(actual_remuneration or 0))
    if compensation <= 0:
        return 0.0
    if compensation <= 5000:
        return 5000.0
    return min(35000.0, float(int((compensation + 499.999999) // 500) * 500))


def get_monthly_contribution_shares(employee, settings, year=None, month=None):
    """Return statutory employee/employer shares using actual monthly bases.

    SSS: 5% employee + 10% employer on MSC, plus employer-only EC.
    PhilHealth: 5% of monthly basic salary within P10,000-P100,000, split 50/50.
    Pag-IBIG: employee 1% if monthly compensation <= P1,500, otherwise 2%;
              employer 2%; contribution basis capped at P5,000.
    """
    if year is None or month is None:
        today = date.today()
        year, month = today.year, today.month

    basis = _monthly_earnings_basis(employee, year, month)
    sss_msc = _sss_msc(basis["sss_remuneration"])

    sss_employee = round(sss_msc * 0.05, 2) if sss_msc else 0.0
    sss_regular_employer = round(sss_msc * 0.10, 2) if sss_msc else 0.0
    sss_ec = 0.0
    if sss_msc:
        sss_ec = 10.0 if sss_msc <= 14500 else 30.0
    sss_employer = round(sss_regular_employer + sss_ec, 2)

    ph_basic = max(0.0, float(basis["philhealth_basic_salary"] or 0))
    if ph_basic <= 0:
        philhealth_employee = philhealth_employer = 0.0
    else:
        ph_base = min(100000.0, max(10000.0, ph_basic))
        philhealth_total = round(ph_base * 0.05, 2)
        # PhilHealth requires the excess centavo, when any, to be borne by
        # the employee; this preserves the exact two-decimal total.
        philhealth_employee = round(philhealth_total / 2, 2)
        philhealth_employer = round(philhealth_total - philhealth_employee, 2)

    pagibig_comp = max(0.0, float(basis["pagibig_monthly_compensation"] or 0))
    if pagibig_comp <= 0:
        pagibig_employee = pagibig_employer = 0.0
    else:
        pagibig_base = min(5000.0, pagibig_comp)
        pagibig_rate = 0.01 if pagibig_base <= 1500 else 0.02
        pagibig_employee = round(pagibig_base * pagibig_rate, 2)
        pagibig_employer = round(pagibig_base * 0.02, 2)

    return {
        "monthly_base": basis["sss_remuneration"],
        "sss_remuneration": basis["sss_remuneration"],
        "philhealth_basic_salary": ph_basic,
        "pagibig_monthly_compensation": pagibig_comp,
        "sss_msc": sss_msc,
        "sss_employee": sss_employee,
        "sss_employer": sss_employer,
        "sss_ec": sss_ec,
        "philhealth_employee": philhealth_employee,
        "philhealth_employer": philhealth_employer,
        "pagibig_employee": pagibig_employee,
        "pagibig_employer": pagibig_employer,
    }


def _withholding_tax_semi_monthly(taxable_compensation):
    """BIR Annex E semi-monthly table, effective Jan 1, 2023 onward."""
    x = max(0.0, float(taxable_compensation or 0))
    if x <= 10417:
        return 0.0
    if x <= 16666:
        return round((x - 10417) * 0.15, 2)
    if x <= 33332:
        return round(937.50 + (x - 16667) * 0.20, 2)
    if x <= 83332:
        return round(4270.70 + (x - 33333) * 0.25, 2)
    if x <= 333332:
        return round(16770.70 + (x - 83333) * 0.30, 2)
    return round(91770.70 + (x - 333333) * 0.35, 2)


def is_second_cutoff(period_start, period_end):
    return (
        period_start.day == 16 and
        period_end == date(
            period_end.year,
            period_end.month,
            calendar.monthrange(period_end.year, period_end.month)[1]
        )
    )


def get_monthly_contribution_record(employee_id, year, month):
    return MonthlyContribution.query.filter_by(
        employee_id=employee_id,
        contribution_year=year,
        contribution_month=month
    ).first()


def upsert_monthly_contribution(employee, period_end, settings, payroll_id=None,
                                collected=None):
    """Create/update the monthly statutory obligation for a completed month."""
    shares = get_monthly_contribution_shares(employee, settings, period_end.year, period_end.month)
    record = get_monthly_contribution_record(employee.id, period_end.year, period_end.month)
    if record is None:
        record = MonthlyContribution(
            employee_id=employee.id,
            contribution_year=period_end.year,
            contribution_month=period_end.month,
        )
        db.session.add(record)

    record.sss_remuneration = round(shares['sss_remuneration'], 2)
    record.sss_msc = round(shares['sss_msc'], 2)
    record.philhealth_basic_salary = round(shares['philhealth_basic_salary'], 2)
    record.pagibig_monthly_compensation = round(shares['pagibig_monthly_compensation'], 2)
    record.sss_due = round(shares['sss_employee'], 2)
    record.philhealth_due = round(shares['philhealth_employee'], 2)
    record.pagibig_due = round(shares['pagibig_employee'], 2)
    record.sss_employer = round(shares['sss_employer'], 2)
    record.philhealth_employer = round(shares['philhealth_employer'], 2)
    record.pagibig_employer = round(shares['pagibig_employer'], 2)
    if payroll_id:
        record.source_payroll_id = payroll_id

    if collected is not None:
        record.sss_collected = round(float(collected.get('sss', 0) or 0), 2)
        record.philhealth_collected = round(float(collected.get('philhealth', 0) or 0), 2)
        record.pagibig_collected = round(float(collected.get('pagibig', 0) or 0), 2)
        record.sss_uncollected = round(max(0, record.sss_due - record.sss_collected), 2)
        record.philhealth_uncollected = round(max(0, record.philhealth_due - record.philhealth_collected), 2)
        record.pagibig_uncollected = round(max(0, record.pagibig_due - record.pagibig_collected), 2)
        record.status = 'Collected' if (
            record.sss_uncollected + record.philhealth_uncollected + record.pagibig_uncollected
        ) <= 0.009 else 'Partially Collected' if (
            record.sss_collected + record.philhealth_collected + record.pagibig_collected
        ) > 0 else 'Uncollected'
    else:
        record.sss_uncollected = record.sss_due
        record.philhealth_uncollected = record.philhealth_due
        record.pagibig_uncollected = record.pagibig_due
        record.status = 'Uncollected'

    return record


def calculate_deductions(employee, gross_pay, settings,
                         period_start, period_end, loan=0,
                         other_deduction=0):
    """Calculate payroll deductions while keeping monthly statutory obligations separate.

    Statutory contributions are monthly obligations and are finalized on the
    16th-to-last-day payroll. The employee can only have the amount actually
    collectible from the current payroll deducted; any remaining employee
    contribution is tracked in MonthlyContribution as uncollected.
    """
    payable_gross = max(0.0, float(gross_pay or 0))
    sss = philhealth = pagibig = withholding_tax = 0.0

    if is_second_cutoff(period_start, period_end):
        shares = get_monthly_contribution_shares(employee, settings, period_end.year, period_end.month)
        statutory_due = round(
            shares['sss_employee'] + shares['philhealth_employee'] + shares['pagibig_employee'], 2
        )
        collectible_statutory = min(statutory_due, payable_gross)

        # Allocate the collectible amount in statutory order so the individual
        # payroll lines always reconcile with the total actually deducted.
        sss = min(shares['sss_employee'], collectible_statutory)
        remaining = collectible_statutory - sss
        philhealth = min(shares['philhealth_employee'], max(0.0, remaining))
        remaining -= philhealth
        pagibig = min(shares['pagibig_employee'], max(0.0, remaining))

    # Withholding tax is a payroll-period tax, not a 2nd-cutoff-only item.
    # The 1st cutoff simply has zero SSS/PhilHealth/Pag-IBIG deduction under
    # the EMS policy, while BIR withholding is still computed for that period.
    attendance_records = payable_attendance_records(employee.id, period_start, period_end)
    hourly_rate = float(employee.daily_rate or 0) / max(float(settings.hours_per_day or 8), 1)
    late_undertime = round(
        sum(float(r.late_hours or 0) + float(r.undertime_hours or 0) for r in attendance_records) * hourly_rate,
        2
    )
    taxable = max(0.0, payable_gross - late_undertime - sss - philhealth - pagibig)
    withholding_tax = _withholding_tax_semi_monthly(taxable) if taxable > 0 else 0.0

    loan = max(0.0, float(loan or 0))
    other_deduction = max(0.0, float(other_deduction or 0))
    total = round(
        sss + philhealth + pagibig + withholding_tax + loan + other_deduction,
        2
    )
    total = min(total, payable_gross)

    return {
        'sss': round(sss, 2),
        'philhealth': round(philhealth, 2),
        'pagibig': round(pagibig, 2),
        'withholding_tax': round(withholding_tax, 2),
        'loan': round(loan, 2),
        'other_deduction': round(other_deduction, 2),
        'cash_advance': 0.0,
        'total_deductions': total,
    }


def calculate_ytd_basic_pay(employee_id, year, through_date=None):
    """Sum basic pay actually covered by payroll records for a calendar year.

    ``through_date`` is used for separated employees so a later payroll period
    cannot accidentally be included after the employee's last day.
    """
    start = date(year, 1, 1)
    end = date(year, 12, 31)
    if through_date is not None:
        end = min(end, through_date)

    records = Payroll.query.filter(
        Payroll.employee_id == employee_id,
        Payroll.period_start >= start,
        Payroll.period_end <= end,
        Payroll.status.in_(["Approved", "Paid"])
    ).all()
    return round(sum(float(p.basic_pay or 0) for p in records), 2)


def calculate_final_pay_values(employee, separation_date, settings):
    """
    Calculate unpaid final-pay coverage, approved OT, and pro-rated 13th month.
    Existing paid/approved payrolls are not duplicated.
    """
    previous_payroll = Payroll.query.filter(
        Payroll.employee_id == employee.id,
        Payroll.period_end < separation_date,
        Payroll.status.in_(["Approved", "Paid"])
    ).order_by(Payroll.period_end.desc()).first()

    coverage_start = (
        previous_payroll.period_end + timedelta(days=1)
        if previous_payroll else (employee.date_hired or separation_date)
    )
    coverage_end = separation_date

    attendance = payable_attendance_records(
        employee.id, coverage_start, coverage_end
    )

    worked_days = len(attendance)
    basic_pay = round(
        max(0.0, float(employee.daily_rate or 0)) * worked_days, 2
    )
    overtime_hours = round(
        sum(float(r.approved_overtime_hours or 0) for r in attendance), 2
    )
    overtime_pay = calculate_overtime_pay(
        employee.daily_rate or 0, overtime_hours, settings
    )

    # 13th month: one-twelfth of basic pay earned in the year.
    ytd_basic = calculate_ytd_basic_pay(employee.id, separation_date.year)
    thirteenth_month = round((ytd_basic + basic_pay) / 12.0, 2)

    gross_pay = round(basic_pay + overtime_pay + thirteenth_month, 2)

    # Only include current-month statutory employee shares when there is no
    # second-cutoff payroll already posted for the separation month.
    second_cutoff = date(
        separation_date.year,
        separation_date.month,
        16
    )
    last_day = calendar.monthrange(separation_date.year, separation_date.month)[1]
    existing_second = Payroll.query.filter(
        Payroll.employee_id == employee.id,
        Payroll.period_start == second_cutoff,
        Payroll.period_end == date(separation_date.year, separation_date.month, last_day),
        Payroll.status.in_(["Approved", "Paid"])
    ).first()

    sss = philhealth = pagibig = withholding_tax = 0.0
    if not existing_second:
        shares = get_monthly_contribution_shares(employee, settings, separation_date.year, separation_date.month)
        sss = shares["sss_employee"]
        philhealth = shares["philhealth_employee"]
        pagibig = shares["pagibig_employee"]
        taxable = max(0.0, gross_pay - sss - philhealth - pagibig)
        withholding_tax = _withholding_tax_semi_monthly(taxable)

    loans, loan_breakdown = get_active_loan_installments(employee.id, separation_date)
    requested_statutory = round(sss + philhealth + pagibig + withholding_tax, 2)
    actual_deductions = round(
        requested_statutory + loan_breakdown["total"], 2
    )
    uncollected = max(0.0, actual_deductions - gross_pay)
    net_pay = round(max(0.0, gross_pay - actual_deductions), 2)

    shares = get_monthly_contribution_shares(employee, settings, separation_date.year, separation_date.month)

    return {
        "last_payroll_end": previous_payroll.period_end if previous_payroll else None,
        "coverage_start": coverage_start,
        "coverage_end": coverage_end,
        "basic_pay": basic_pay,
        "overtime_hours": overtime_hours,
        "overtime_pay": overtime_pay,
        "thirteenth_month": thirteenth_month,
        "gross_pay": gross_pay,
        "sss_employee": round(sss, 2),
        "philhealth_employee": round(philhealth, 2),
        "pagibig_employee": round(pagibig, 2),
        "withholding_tax": round(withholding_tax, 2),
        "requested_statutory": requested_statutory,
        "actual_deductions": actual_deductions,
        "uncollected_deductions": round(uncollected, 2),
        "net_pay": net_pay,
        "sss_employer": shares["sss_employer"],
        "philhealth_employer": shares["philhealth_employer"],
        "pagibig_employer": shares["pagibig_employer"],
    }


def calculate_payroll_values(employee, period_start, period_end, settings,
                             adjustment_addition=0,
                             adjustment_deduction=0,
                             adjustment_note=None):
    """
    Payroll is generated for one employee using:
      - actual regular paid hours
      - daily allowance x worked days
      - daily incentive x worked days
      - employee-specific loan deduction
      - optional one-time payroll adjustment
    """
    basic_pay = round(
        calculate_basic_pay(employee, period_start, period_end, settings), 2
    )

    holiday_values = calculate_holiday_earnings(
        employee, period_start, period_end, settings
    )
    overtime_hours = round(
        holiday_values["ordinary_ot_hours"] + sum(x[4] for x in holiday_values["holiday_days"]), 2
    )
    overtime_pay = round(
        holiday_values["ordinary_ot_pay"] + holiday_values["holiday_ot_pay"] + holiday_values["rest_day_ot_pay"], 2
    )

    worked_days = get_worked_days(
        employee.id, period_start, period_end
    )
    leave_days = leave_days_for_period(employee, period_start, period_end)
    vl_days = leave_days["vl_days"]
    sl_days = leave_days["sl_days"]
    daily_rate = max(0.0, float(employee.daily_rate or 0))
    vl_pay = round(daily_rate * vl_days, 2)
    sl_pay = round(daily_rate * sl_days, 2)
    leave_conversion = leave_conversion_for_period(employee, period_start, period_end)

    allowance = round(
        max(0.0, float(employee.daily_allowance or 0)) * worked_days,
        2
    )
    incentive = round(
        max(0.0, float(employee.daily_incentive or 0)) * worked_days,
        2
    )

    adjustment_addition = round(
        float(adjustment_addition or 0), 2
    )
    adjustment_deduction = round(
        max(0.0, float(adjustment_deduction or 0)), 2
    )

    _, loan_breakdown = get_active_loan_installments(employee.id, period_end)
    loan = loan_breakdown["total"]

    attendance_records = payable_attendance_records(
        employee.id, period_start, period_end
    )

    worked_days = len(attendance_records)
    regular_hours = round(sum(float(r.total_hours or 0) for r in attendance_records), 2)
    late_hours = round(sum(float(r.late_hours or 0) for r in attendance_records), 2)
    undertime_hours = round(sum(float(r.undertime_hours or 0) for r in attendance_records), 2)
    hourly_rate = float(employee.daily_rate or 0) / max(float(settings.hours_per_day or 8), 1)

    late_deduction = round(
        late_hours * hourly_rate,
        2
    )
    undertime_deduction = round(
        undertime_hours * hourly_rate,
        2
    )

    # Gross earnings are based on completed worked days and approved earnings.
    # Late/undertime are shown separately as employee deductions, so they are
    # not silently buried inside Basic Pay and are not double-counted.
    gross_pay = round(
        basic_pay + holiday_values["holiday_premium"] + holiday_values["rest_day_premium"] + overtime_pay + allowance + incentive +
        leave_conversion["vl_conversion"] + leave_conversion["sl_conversion"] + adjustment_addition,
        2
    )

    deductions = calculate_deductions(
        employee,
        gross_pay,
        settings,
        period_start,
        period_end,
        loan,
        adjustment_deduction
    )

    # Late + undertime are employee deductions too.
    deductions["late_deduction"] = late_deduction
    deductions["undertime_deduction"] = undertime_deduction
    deductions["total_deductions"] = round(
        float(deductions.get("total_deductions", 0))
        + late_deduction + undertime_deduction,
        2
    )

    deductions["total_deductions"] = min(
        round(float(deductions["total_deductions"] or 0), 2),
        max(0.0, float(gross_pay or 0))
    )
    net_pay = round(
        max(0.0, gross_pay - deductions["total_deductions"]),
        2
    )

    return {
        "basic_pay": basic_pay,
        "worked_days": worked_days,
        "regular_hours": regular_hours,
        "late_hours": late_hours,
        "undertime_hours": undertime_hours,
        "undertime_deduction": undertime_deduction,
        "late_deduction": late_deduction,
        "overtime_hours": overtime_hours,
        "overtime_pay": overtime_pay,
        "regular_ot_hours": holiday_values["ordinary_ot_hours"],
        "regular_ot_pay": holiday_values["ordinary_ot_pay"],
        "holiday_ot_hours": holiday_values["holiday_ot_hours"],
        "holiday_ot_pay": holiday_values["holiday_ot_pay"],
        "rest_day_ot_hours": holiday_values["rest_day_ot_hours"],
        "rest_day_ot_pay": holiday_values["rest_day_ot_pay"],
        "holiday_pay": holiday_values["holiday_premium"],
        "rest_day_premium": holiday_values["rest_day_premium"],
        "absent_days": get_attendance_summary(employee.id, period_start, period_end)["absent_days"],
        "leave_with_pay_days": get_attendance_summary(employee.id, period_start, period_end)["leave_with_pay_days"],
        "leave_without_pay_days": get_attendance_summary(employee.id, period_start, period_end)["leave_without_pay_days"],
        "vl_days": vl_days,
        "sl_days": sl_days,
        "vl_pay": vl_pay,
        "sl_pay": sl_pay,
        "vl_conversion": leave_conversion["vl_conversion"],
        "sl_conversion": leave_conversion["sl_conversion"],
        "allowance": allowance,
        "incentive": incentive,
        "gross_pay": gross_pay,
        "loan_sss_salary": loan_breakdown["loan_sss_salary"],
        "loan_sss_calamity": loan_breakdown["loan_sss_calamity"],
        "loan_pagibig_mpl": loan_breakdown["loan_pagibig_mpl"],
        "loan_pagibig_calamity": loan_breakdown["loan_pagibig_calamity"],
        **deductions,
        "adjustment_addition": adjustment_addition,
        "adjustment_deduction": adjustment_deduction,
        "adjustment_note": adjustment_note,
        "net_pay": net_pay,
        "status": "Approved",
    }


def _json_load(value, default=None):
    if default is None:
        default = []
    try:
        data = json.loads(value or "")
        return data if isinstance(data, type(default)) else default
    except (TypeError, ValueError, json.JSONDecodeError):
        return default


def _form_json(name, default=None):
    raw = request.form.get(name, "").strip()
    if not raw:
        return default if default is not None else []
    try:
        data = json.loads(raw)
        return data
    except (TypeError, ValueError, json.JSONDecodeError):
        return default if default is not None else []


def save_employee_form(employee):
    employee.employee_id = request.form.get("employee_id", "").strip()
    employee.first_name = request.form.get("first_name", "").strip()
    employee.middle_name = request.form.get("middle_name", "").strip() or None
    employee.last_name = request.form.get("last_name", "").strip()
    employee.nickname = request.form.get("nickname", "").strip() or None
    employee.email = request.form.get("email", "").strip() or None
    employee.company_email = request.form.get("company_email", "").strip() or None
    employee.phone = request.form.get("phone", "").strip() or None
    employee.secondary_phone = request.form.get("secondary_phone", "").strip() or None
    employee.address = request.form.get("address", "").strip() or None
    employee.permanent_address = request.form.get("permanent_address", "").strip() or None
    employee.social_media = request.form.get("social_media", "").strip() or None
    employee.department = request.form.get("department", "").strip() or None
    employee.position = request.form.get("position", "").strip() or None
    employee.reporting_to = request.form.get("reporting_to", "").strip() or None
    employee.date_hired = parse_date(request.form.get("date_hired"))
    employee.regularization_date = parse_date(request.form.get("regularization_date"))
    employee.date_of_birth = request.form.get("date_of_birth", "").strip() or None
    employee.place_of_birth = request.form.get("place_of_birth", "").strip() or None
    employee.sex = request.form.get("sex", "").strip() or None
    employee.civil_status = request.form.get("civil_status", "").strip() or None
    employee.nationality = request.form.get("nationality", "").strip() or None
    employee.religion = request.form.get("religion", "").strip() or None
    employee.blood_type = request.form.get("blood_type", "").strip() or None
    employee.height = request.form.get("height", "").strip() or None
    employee.weight = request.form.get("weight", "").strip() or None
    employee.work_location = request.form.get("work_location", "").strip() or None
    employee.employment_type = request.form.get("employment_type", "Full-Time").strip() or "Full-Time"
    employee.schedule_type = request.form.get("schedule_type", "Fixed").strip() or "Fixed"
    employee.schedule_details = request.form.get("schedule_details", "").strip() or None
    employee.professional_summary = request.form.get("professional_summary", "").strip() or None
    employee.duties = request.form.get("duties", "").strip() or None

    employee.vl_entitlement = max(0.0, float(request.form.get("vl_entitlement") or 0))
    employee.sl_entitlement = max(0.0, float(request.form.get("sl_entitlement") or 0))
    employee.employment_status = request.form.get("employment_status", "Active")
    employee.daily_rate = float(request.form.get("daily_rate") or 0)
    employee.daily_allowance = max(0.0, float(request.form.get("daily_allowance") or 0))
    employee.daily_incentive = max(0.0, float(request.form.get("daily_incentive") or 0))
    employee.loan_deduction = max(0.0, float(request.form.get("loan_deduction") or 0))
    employee.separation_date = parse_date(request.form.get("separation_date"))
    employee.separation_type = request.form.get("separation_type", "").strip() or None
    employee.separation_reason = request.form.get("separation_reason", "").strip() or None

    employee.sss_number = request.form.get("sss_number", "").strip() or None
    employee.philhealth_number = request.form.get("philhealth_number", "").strip() or None
    employee.pagibig_number = request.form.get("pagibig_number", "").strip() or None
    employee.tin_number = request.form.get("tin_number", "").strip() or None
    employee.umid_number = request.form.get("umid_number", "").strip() or None
    employee.national_id = request.form.get("national_id", "").strip() or None
    employee.voters_id = request.form.get("voters_id", "").strip() or None
    employee.drivers_license = request.form.get("drivers_license", "").strip() or None
    employee.passport_number = request.form.get("passport_number", "").strip() or None
    employee.passport_expiry = request.form.get("passport_expiry", "").strip() or None

    employee.bank_name = request.form.get("bank_name", "").strip() or None
    employee.bank_account_name = request.form.get("bank_account_name", "").strip() or None
    employee.bank_account_number = request.form.get("bank_account_number", "").strip() or None

    employee.emergency_name = request.form.get("emergency_name", "").strip() or None
    employee.emergency_relationship = request.form.get("emergency_relationship", "").strip() or None
    employee.emergency_phone = request.form.get("emergency_phone", "").strip() or None
    employee.emergency_address = request.form.get("emergency_address", "").strip() or None

    for attr, form_name in {
        "education_json": "education_json",
        "work_experience_json": "work_experience_json",
        "skills_json": "skills_json",
        "family_json": "family_json",
        "references_json": "references_json",
        "documents_json": "documents_json",
    }.items():
        data = _form_json(form_name, [])
        setattr(employee, attr, json.dumps(data, ensure_ascii=False))

def save_profile_photo(employee, file):
    if not file or not file.filename:
        return

    extension = file.filename.rsplit(".", 1)[-1].lower()
    if extension not in ALLOWED_IMAGE_EXTENSIONS:
        raise ValueError("Only JPG, JPEG, PNG, or WEBP images are allowed.")

    filename = secure_filename(
        f"employee_{employee.id}_{int(datetime.utcnow().timestamp())}.{extension}"
    )
    path = os.path.join(PROFILE_DIR, filename)
    file.save(path)
    employee.profile_photo = filename


@app.route("/", methods=["GET", "POST"])
def login():
    device_id = get_device_id()

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        user = User.query.filter_by(username=username).first()

        if not user or not user.check_password(password):
            flash("Invalid username or password.", "danger")
            return render_template("login.html", device_id=device_id)

        # Admin can bootstrap/register the first admin browser by logging in.
        # All employee/admin accounts must use an approved company device.
        if user.role == "admin":
            if not device_is_allowed(user):
                register_or_update_device(
                    device_id,
                    user=user,
                    device_name="Admin Device",
                    status="Active"
                )
                db.session.commit()
                flash("This admin device has been registered successfully.", "success")
        else:
            if not device_is_allowed(user):
                pending = create_pending_device_for_user(user)
                flash(
                    f"This device is not approved yet. Device ID: {pending.device_id}. "
                    "Please give this ID to your Admin for approval.",
                    "warning"
                )
                return render_template(
                    "device_blocked.html",
                    device=pending,
                    employee=user.employee
                )

        session.clear()
        session["user_id"] = user.id
        session["role"] = user.role
        session["device_id"] = device_id

        if user.role == "admin":
            return redirect(url_for("dashboard"))
        return redirect(url_for("employee_dashboard"))

    return render_template("login.html", device_id=device_id)


@app.route("/device-status")
def device_status():
    """Show the current browser's Device ID and registration status."""
    device = current_device()
    return render_template("device_status.html", device=device, device_id=get_device_id())


@app.route("/device-register", methods=["GET", "POST"])
def device_register():
    """Admin-only bootstrap/registration helper for the current browser."""
    device_id = get_device_id()
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username, role="admin").first()

        if not user or not user.check_password(password):
            flash("Invalid admin credentials.", "danger")
            return render_template("device_register.html", device_id=device_id)

        register_or_update_device(
            device_id,
            user=user,
            device_name=request.form.get("device_name", "").strip() or "Admin Device",
            status="Active"
        )
        db.session.commit()
        flash("Device registered successfully. You may now log in.", "success")
        return redirect(url_for("login"))

    return render_template("device_register.html", device_id=device_id)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/account", methods=["GET", "POST"])
def account():
    """View the current account and securely change its password."""
    if not session.get("user_id"):
        return redirect(url_for("login"))

    user = User.query.get_or_404(session["user_id"])

    if request.method == "POST":
        current_password = request.form.get("current_password", "")
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")

        if not user.check_password(current_password):
            flash("Current password is incorrect.", "danger")
            return render_template("account.html", user=user)

        if len(new_password) < 8:
            flash("New password must be at least 8 characters.", "danger")
            return render_template("account.html", user=user)

        if new_password != confirm_password:
            flash("New password and confirmation do not match.", "danger")
            return render_template("account.html", user=user)

        if new_password == current_password:
            flash("New password must be different from the current password.", "danger")
            return render_template("account.html", user=user)

        user.set_password(new_password)
        db.session.commit()

        # Force a fresh login after a password change.
        session.clear()
        flash("Password changed successfully. Please log in again.", "success")
        return redirect(url_for("login"))

    return render_template("account.html", user=user)


@app.route("/admin/devices", methods=["GET", "POST"])
@admin_required
def admin_devices():
    if request.method == "POST":
        action = request.form.get("action", "").strip()
        device_id = request.form.get("device_id", "").strip().upper()
        device = Device.query.filter_by(device_id=device_id).first()

        if action == "approve":
            if not device:
                flash("Device ID was not found. Ask the employee to open EMS and submit the displayed Device ID.", "danger")
            else:
                device.status = "Active"
                device.device_name = request.form.get("device_name", "").strip() or device.device_name
                # Company-device whitelist: never bind a device to one employee.
                device.employee_id = None
                device.approved_at = datetime.utcnow()
                db.session.commit()
                flash("Device approved.", "success")

        elif action == "block":
            if device:
                device.status = "Blocked"
                db.session.commit()
                flash("Device blocked.", "success")
            else:
                flash("Device not found.", "danger")

        elif action == "delete":
            if device:
                db.session.delete(device)
                db.session.commit()
                flash("Device registration deleted. The device will be treated as unregistered.", "success")
            else:
                flash("Device not found.", "danger")

        return redirect(url_for("admin_devices"))

    devices = Device.query.order_by(Device.status.asc(), Device.created_at.desc()).all()
    employees = Employee.query.order_by(Employee.last_name.asc(), Employee.first_name.asc()).all()
    return render_template(
        "admin_devices.html",
        devices=devices,
        employees=employees,
        current_device_id=get_device_id()
    )


def _parse_commission_rate(value):
    """Accept Excel values such as 10%, 0.10, or 10 and return decimal rate."""
    if value is None or str(value).strip() == "":
        raise ValueError("Commission Rate is required")
    if isinstance(value, str):
        raw = value.strip().replace(",", "")
        if raw.endswith("%"):
            return float(raw[:-1]) / 100.0
        number = float(raw)
    else:
        number = float(value)
    # Excel percentage cells normally arrive as 0.10; plain 10 means 10%.
    return number / 100.0 if number > 1 else number


def _money(value):
    return round(float(value or 0), 2)



def _valid_second_cutoff_payroll_query(year=None, month=None):
    """Return second-cutoff payrolls that are the source of monthly contributions."""
    query = Payroll.query.filter(Payroll.status.in_(["Approved", "Paid"]))
    if year is not None and month is not None:
        start = date(year, month, 16)
        end = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
        query = query.filter(Payroll.period_start == start, Payroll.period_end == date(year, month, calendar.monthrange(year, month)[1]))
    return query


def cleanup_orphan_financial_records(year=None, month=None):
    """Remove financial rows that no longer have a valid source payroll.

    Payroll is the source of truth. A monthly contribution is valid only when
    it points to an existing Approved/Paid 2nd-cutoff payroll for the same
    employee and month. This also repairs legacy rows created before
    source_payroll_id existed.

    This function is intentionally defensive because production databases can
    contain old financial rows from previous EMS versions.
    """
    # Salary expenses must point to an existing payroll.
    salary_expenses = Expense.query.filter_by(source_type="payroll_salary").all()
    for expense in salary_expenses:
        try:
            payroll_id = int((expense.source_key or "").split(":", 1)[1])
        except (ValueError, IndexError, TypeError):
            payroll_id = None
        if payroll_id is None or db.session.get(Payroll, payroll_id) is None:
            db.session.delete(expense)

    # Every contribution record must have a live 2nd-cutoff payroll. Do not
    # keep legacy rows merely because another payroll happens to share a month.
    contrib_query = MonthlyContribution.query
    if year is not None and month is not None:
        contrib_query = contrib_query.filter_by(contribution_year=year, contribution_month=month)
    for record in contrib_query.all():
        source = db.session.get(Payroll, record.source_payroll_id) if record.source_payroll_id else None
        valid = bool(source and is_second_cutoff(source.period_start, source.period_end)
                     and source.status in ("Approved", "Paid")
                     and source.employee_id == record.employee_id
                     and source.period_end.year == record.contribution_year
                     and source.period_end.month == record.contribution_month)
        if not valid:
            db.session.delete(record)

    # Contribution expense rows are monthly employer-cost aggregates. Remove
    # them whenever no valid contribution records remain for that month.
    contribution_expenses = Expense.query.filter_by(source_type="payroll_contribution").all()
    for expense in contribution_expenses:
        try:
            period_key, field = (expense.source_key or "").split(":", 1)
            y, m = [int(x) for x in period_key.split("-")]
        except (ValueError, TypeError):
            db.session.delete(expense)
            continue
        if year is not None and month is not None and (y != year or m != month):
            continue
        valid_exists = False
        for record in MonthlyContribution.query.filter_by(contribution_year=y, contribution_month=m).all():
            source = db.session.get(Payroll, record.source_payroll_id) if record.source_payroll_id else None
            if (source and is_second_cutoff(source.period_start, source.period_end)
                    and source.status in ("Approved", "Paid")
                    and source.employee_id == record.employee_id):
                valid_exists = True
                break
        if not valid_exists:
            db.session.delete(expense)

    db.session.commit()


def sync_payroll_salary_expenses(year=None, month=None):
    """Create/update salary expense rows from approved/paid payrolls.

    Approved = Pending, Paid = Paid. Deleted payrolls are removed by the
    orphan cleanup above, so salary cannot remain stale in the reports.
    """
    cleanup_orphan_financial_records(year, month)
    query = Payroll.query.filter(Payroll.status.in_(["Approved", "Paid"]))
    if year is not None and month is not None:
        start = date(year, month, 1)
        end = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
        query = query.filter(Payroll.period_end >= start, Payroll.period_end < end)
    rows = query.all()
    for payroll in rows:
        period_key = f"{payroll.period_start.isoformat()}:{payroll.period_end.isoformat()}"
        source_key = f"payroll_salary:{payroll.id}"
        employee_name = payroll.employee.full_name if getattr(payroll.employee, "full_name", None) else f"{payroll.employee.first_name} {payroll.employee.last_name}".strip()
        amount = round(float(payroll.net_pay or 0), 2)
        if amount <= 0:
            continue
        status = "Paid" if payroll.status == "Paid" else "Pending"
        expense = Expense.query.filter_by(source_type="payroll_salary", source_key=source_key).first()
        remarks = f"Auto-generated from Payroll {period_key}. Employee: {employee_name}. Payroll status: {payroll.status}."
        if not expense:
            db.session.add(Expense(
                expense_date=payroll.pay_date or payroll.period_end,
                name=f"Salary - {employee_name} ({payroll.period_start:%b %d}–{payroll.period_end:%b %d, %Y})",
                category="Salary / Payroll", amount=amount,
                due_date=payroll.pay_date or payroll.period_end,
                paid_date=(payroll.pay_date if payroll.status == "Paid" else None),
                status=status, remarks=remarks, source_type="payroll_salary", source_key=source_key
            ))
        else:
            expense.amount = amount
            expense.name = f"Salary - {employee_name} ({payroll.period_start:%b %d}–{payroll.period_end:%b %d, %Y})"
            expense.due_date = payroll.pay_date or payroll.period_end
            expense.status = status
            expense.paid_date = payroll.pay_date if payroll.status == "Paid" else None
            expense.remarks = remarks
    db.session.commit()


@app.route("/expenses", methods=["GET", "POST"])
@admin_required
def expense_report():
    """Separate expense trackers. Only Paid expenses flow into the financial summary."""
    today = date.today()
    try:
        year = int(request.args.get("year", today.year))
        month = int(request.args.get("month", today.month))
    except (TypeError, ValueError):
        year, month = today.year, today.month
    month = month if 1 <= month <= 12 else today.month

    if request.method == "POST":
        action = request.form.get("action", "add")
        if action == "sync_contributions":
            sync_contribution_expenses(year, month)
            flash("Mandatory contribution obligations synchronized from Payroll.", "success")
            return redirect(url_for("expense_report", year=year, month=month, view="contributions"))
        if action == "update_payment_meta":
            expense = Expense.query.get_or_404(int(request.form["expense_id"]))
            if expense.source_type != "payroll_contribution":
                flash("Payment details can only be updated for government contribution tracker rows.", "danger")
                return redirect(url_for("expense_report", year=year, month=month, view="contributions"))
            due_raw = request.form.get("due_date", "").strip()
            expense.due_date = datetime.strptime(due_raw, "%Y-%m-%d").date() if due_raw else None
            expense.payment_reference = request.form.get("payment_reference", "").strip() or None
            db.session.commit()
            flash("Government payment details updated.", "success")
            return redirect(url_for("expense_report", year=year, month=month, view="contributions"))

        if action in {"mark_paid", "mark_pending"}:
            expense = Expense.query.get_or_404(int(request.form["expense_id"]))
            if expense.source_type == "payroll_contribution" and action == "mark_pending":
                expense.status = "Pending"
                expense.paid_date = None
            elif action == "mark_paid":
                expense.status = "Paid"
                paid_raw = request.form.get("paid_date", "").strip()
                expense.paid_date = datetime.strptime(paid_raw, "%Y-%m-%d").date() if paid_raw else today
            else:
                expense.status = "Pending"
                expense.paid_date = None
            db.session.commit()
            flash("Expense status updated.", "success")
            return redirect(url_for("expense_report", year=year, month=month, view=request.args.get("view", "other")))

        try:
            expense_date = datetime.strptime(request.form["expense_date"], "%Y-%m-%d").date()
            amount = float(request.form["amount"])
            if amount < 0:
                raise ValueError
            due_raw = request.form.get("due_date", "").strip()
            due_date = datetime.strptime(due_raw, "%Y-%m-%d").date() if due_raw else None
            category = request.form.get("category", "Other Expense").strip() or "Other Expense"
            expense = Expense(
                expense_date=expense_date,
                name=request.form["name"].strip(),
                category=category,
                amount=amount,
                due_date=due_date,
                status=request.form.get("status", "Pending"),
                remarks=request.form.get("remarks", "").strip(),
                source_type="manual",
            )
            db.session.add(expense)
            db.session.commit()
            flash("Expense added.", "success")
        except (KeyError, ValueError):
            flash("Please check the expense details and amount.", "danger")
        return redirect(url_for("expense_report", year=year, month=month, view=request.form.get("view", "other")))

    sync_contribution_expenses(year, month)
    sync_payroll_salary_expenses(year, month)
    # Keep 13th Month Tracker and financial expense records synchronized even
    # when the admin opens Expense Trackers directly.
    sync_thirteenth_month_records(year)
    sync_thirteenth_month_expenses(year)
    start = date(year, month, 1)
    end = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
    view = request.args.get("view", "other")
    if view == "contributions":
        expenses = Expense.query.filter(Expense.expense_date >= start, Expense.expense_date < end,
                                        Expense.source_type == "payroll_contribution").order_by(Expense.due_date.asc().nullslast(), Expense.id.desc()).all()
    elif view == "tax":
        expenses = Expense.query.filter(Expense.expense_date >= start, Expense.expense_date < end,
                                        Expense.category == "Tax").order_by(Expense.status.asc(), Expense.due_date.asc().nullslast(), Expense.id.desc()).all()
    else:
        expenses = Expense.query.filter(Expense.expense_date >= start, Expense.expense_date < end,
                                        Expense.category.notin_(["Tax", "Government"]),
                                        Expense.source_type != "payroll_contribution").order_by(Expense.status.asc(), Expense.due_date.asc().nullslast(), Expense.id.desc()).all()

    all_expenses = Expense.query.filter(Expense.expense_date >= start, Expense.expense_date < end).all()
    paid = round(sum(float(e.amount or 0) for e in all_expenses if e.status == "Paid"), 2)
    pending = round(sum(float(e.amount or 0) for e in all_expenses if e.status == "Pending"), 2)
    overdue = round(sum(float(e.amount or 0) for e in all_expenses if e.status == "Pending" and e.due_date and e.due_date < today), 2)

    contribution_records = (MonthlyContribution.query
        .join(Payroll, MonthlyContribution.source_payroll_id == Payroll.id)
        .filter(
            MonthlyContribution.contribution_year == year,
            MonthlyContribution.contribution_month == month,
            Payroll.status.in_(["Approved", "Paid"]),
            Payroll.period_start == date(year, month, 16),
            Payroll.period_end == date(year, month, calendar.monthrange(year, month)[1]),
        ).all())
    contribution_payment_rows = []
    for label, field in (("SSS", "sss"), ("PhilHealth", "philhealth"), ("Pag-IBIG", "pagibig")):
        employee_due = round(sum(float(getattr(r, f"{field}_due") or 0) for r in contribution_records), 2)
        employee_collected = round(sum(float(getattr(r, f"{field}_collected") or 0) for r in contribution_records), 2)
        company_share = round(sum(float(getattr(r, f"{field}_employer") or 0) for r in contribution_records), 2)
        total_due = round(employee_due + company_share, 2)
        source_key = f"{year}-{month:02d}:{field}"
        expense_row = Expense.query.filter_by(source_type="payroll_contribution", source_key=source_key).first()
        if total_due > 0:
            contribution_payment_rows.append({
                "label": label,
                "employee_due": employee_due,
                "employee_collected": employee_collected,
                "company_share": company_share,
                "total_due": total_due,
                "employee_count": len({r.employee_id for r in contribution_records}),
                "expense": expense_row,
                "status": expense_row.status if expense_row else "Pending",
                "due_date": expense_row.due_date if expense_row else None,
                "paid_date": expense_row.paid_date if expense_row else None,
                "payment_reference": expense_row.payment_reference if expense_row else None,
            })

    if view == "contributions":
        paid = round(sum(r["total_due"] for r in contribution_payment_rows if r["status"] == "Paid"), 2)
        pending = round(sum(r["total_due"] for r in contribution_payment_rows if r["status"] == "Pending"), 2)
        overdue = round(sum(r["total_due"] for r in contribution_payment_rows
                            if r["status"] == "Pending" and r["due_date"] and r["due_date"] < today), 2)

    employee_remittance = round(sum(r["employee_collected"] for r in contribution_payment_rows), 2)
    employer_cost = round(sum(r["company_share"] for r in contribution_payment_rows), 2)
    return render_template("expense_report.html", expenses=expenses, year=year, month=month, view=view,
                           paid=paid, pending=pending, overdue=overdue,
                           employee_remittance=employee_remittance, employer_cost=employer_cost,
                           contribution_payment_rows=contribution_payment_rows, today=today)


@app.route("/expenses/government/<agency>/employees")
@admin_required
def government_contribution_employees(agency):
    """Show employee-level contribution details for one government agency."""
    agency_map = {"sss": ("SSS", "sss"), "philhealth": ("PhilHealth", "philhealth"), "pagibig": ("Pag-IBIG", "pagibig")}
    if agency.lower() not in agency_map:
        abort(404)
    label, field = agency_map[agency.lower()]
    today = date.today()
    try:
        year = int(request.args.get("year", today.year))
        month = int(request.args.get("month", today.month))
    except (TypeError, ValueError):
        year, month = today.year, today.month
    if month < 1 or month > 12:
        month = today.month

    purge_month_without_payroll(year, month)
    cleanup_orphan_financial_records(year, month)
    records = (MonthlyContribution.query
               .join(Payroll, MonthlyContribution.source_payroll_id == Payroll.id)
               .join(Employee)
               .filter(
                   MonthlyContribution.contribution_year == year,
                   MonthlyContribution.contribution_month == month,
                   Payroll.status.in_(["Approved", "Paid"]),
                   Payroll.period_start == date(year, month, 16),
                   Payroll.period_end == date(year, month, calendar.monthrange(year, month)[1]),
               )
               .order_by(Employee.last_name.asc(), Employee.first_name.asc()).all())

    rows = []
    for r in records:
        ee = float(getattr(r, f"{field}_due") or 0)
        collected = float(getattr(r, f"{field}_collected") or 0)
        er = float(getattr(r, f"{field}_employer") or 0)
        rows.append({
            "employee": r.employee,
            "employee_share": round(ee, 2),
            "employee_collected": round(collected, 2),
            "company_share": round(er, 2),
            "total": round(ee + er, 2),
            "payroll_id": r.source_payroll_id,
        })
    totals = {
        "employee_share": round(sum(x["employee_share"] for x in rows), 2),
        "employee_collected": round(sum(x["employee_collected"] for x in rows), 2),
        "company_share": round(sum(x["company_share"] for x in rows), 2),
    }
    totals["total"] = round(totals["employee_share"] + totals["company_share"], 2)
    return render_template("government_contribution_employees.html", agency=label,
                           field=field, rows=rows, totals=totals, year=year, month=month)


@app.route("/expenses/government/<agency>/export")
@admin_required
def export_government_contribution(agency):
    """Export employee-level government contribution list as an Excel workbook."""
    agency_map = {"sss": ("SSS", "sss"), "philhealth": ("PhilHealth", "philhealth"), "pagibig": ("Pag-IBIG", "pagibig")}
    if agency.lower() not in agency_map:
        abort(404)
    label, field = agency_map[agency.lower()]
    today = date.today()
    try:
        year = int(request.args.get("year", today.year))
        month = int(request.args.get("month", today.month))
    except (TypeError, ValueError):
        year, month = today.year, today.month
    if month < 1 or month > 12:
        month = today.month

    purge_month_without_payroll(year, month)
    cleanup_orphan_financial_records(year, month)
    records = (MonthlyContribution.query
               .join(Payroll, MonthlyContribution.source_payroll_id == Payroll.id)
               .join(Employee)
               .filter(
                   MonthlyContribution.contribution_year == year,
                   MonthlyContribution.contribution_month == month,
                   Payroll.status.in_(["Approved", "Paid"]),
                   Payroll.period_start == date(year, month, 16),
                   Payroll.period_end == date(year, month, calendar.monthrange(year, month)[1]),
               )
               .order_by(Employee.last_name.asc(), Employee.first_name.asc()).all())

    wb = Workbook()
    ws = wb.active
    ws.title = label[:31]
    headers = ["Employee Name", "Employee ID", "Government ID", "Employee Share", "Employee Collected", "Company Share", "Total Remittance"]
    ws.append([f"{label} Contribution List - {month:02d}/{year}"])
    ws.append([])
    ws.append(headers)
    for cell in ws[3]:
        cell.font = Font(bold=True)
    employee_total = collected_total = company_total = 0.0
    for r in records:
        emp = r.employee
        name = getattr(emp, "full_name", None) or f"{emp.first_name} {emp.last_name}".strip()
        employee_id = getattr(emp, "employee_id", None) or getattr(emp, "employee_code", None) or ""
        gov_id = ""
        if field == "sss":
            gov_id = getattr(emp, "sss_number", None) or ""
        elif field == "philhealth":
            gov_id = getattr(emp, "philhealth_number", None) or ""
        elif field == "pagibig":
            gov_id = getattr(emp, "pagibig_number", None) or ""
        ee = float(getattr(r, f"{field}_due") or 0)
        collected = float(getattr(r, f"{field}_collected") or 0)
        er = float(getattr(r, f"{field}_employer") or 0)
        ws.append([name, employee_id, gov_id, ee, collected, er, ee + er])
        employee_total += ee; collected_total += collected; company_total += er
    ws.append([])
    ws.append(["TOTAL", "", "", employee_total, collected_total, company_total, employee_total + company_total])
    for row in ws.iter_rows(min_row=4, min_col=4, max_col=7):
        for cell in row:
            cell.number_format = '₱#,##0.00'
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    widths = [30, 18, 22, 18, 20, 18, 20]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A4"
    out = BytesIO(); wb.save(out); out.seek(0)
    filename = secure_filename(f"{label.lower().replace('-', '_').replace(' ', '_')}_contribution_list_{year}_{month:02d}.xlsx")
    return send_file(out, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def sync_contribution_expenses(year, month):
    """Create/update statutory remittance checklist rows from active payrolls."""
    cleanup_orphan_financial_records(year, month)
    rows = (MonthlyContribution.query
            .join(Payroll, MonthlyContribution.source_payroll_id == Payroll.id)
            .filter(
                MonthlyContribution.contribution_year == year,
                MonthlyContribution.contribution_month == month,
                Payroll.status.in_(["Approved", "Paid"]),
                Payroll.period_start == date(year, month, 16),
                Payroll.period_end == date(year, month, calendar.monthrange(year, month)[1]),
            ).all())
    period_key = f"{year}-{month:02d}"
    specs = [("SSS", "sss"), ("PhilHealth", "philhealth"), ("Pag-IBIG", "pagibig")]
    period_date = date(year, month, 1)
    for label, field in specs:
        # The company expense is employer share only. Employee deductions are
        # withheld/remitted and are shown separately as employee obligations.
        employer_total = round(sum(float(getattr(r, f"{field}_employer") or 0) for r in rows), 2)
        source_key = f"{period_key}:{field}"
        expense = Expense.query.filter_by(source_type="payroll_contribution", source_key=source_key).first()
        if employer_total <= 0:
            if expense is not None and expense.status != "Paid":
                db.session.delete(expense)
            continue
        employee_due = round(sum(float(getattr(r, f"{field}_due") or 0) for r in rows), 2)
        employee_collected = round(sum(float(getattr(r, f"{field}_collected") or 0) for r in rows), 2)
        remarks = (
            f"Auto-generated from Payroll. Employee share due: ₱{employee_due:,.2f}; "
            f"actually deducted: ₱{employee_collected:,.2f}; Company share: ₱{employer_total:,.2f}."
        )
        if not expense:
            db.session.add(Expense(
                expense_date=period_date, name=f"{label} - {period_key}", category="Government",
                amount=employer_total, status="Pending", source_type="payroll_contribution",
                source_key=source_key, remarks=remarks
            ))
        elif expense.status != "Paid":
            expense.amount = employer_total
            expense.remarks = remarks
    db.session.commit()


def sync_contribution_expenses_all_months():
    """Sync all payroll contribution periods into the expense checklist."""
    periods = db.session.query(
        MonthlyContribution.contribution_year, MonthlyContribution.contribution_month
    ).distinct().all()
    for year, month in periods:
        sync_contribution_expenses(year, month)


@app.route("/financial-summary", methods=["GET"])
@admin_required
def financial_summary():
    today = date.today()
    try:
        year = int(request.args.get("year", today.year))
        month = int(request.args.get("month", today.month))
    except (TypeError, ValueError):
        year, month = today.year, today.month
    if month < 1 or month > 12:
        month = today.month

    start = date(year, month, 1)
    end = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)

    # Production-safe financial synchronization. Ensure legacy databases have
    # the V80 source-link columns before any ORM query uses them, then enforce
    # Payroll as the single source of truth.
    migrate_v80_financial_link_columns()
    purge_month_without_payroll(year, month)
    cleanup_orphan_financial_records(year, month)
    sync_contribution_expenses(year, month)
    sync_payroll_salary_expenses(year, month)
    # 13th Month is a real company expense source. Synchronize it before the
    # monthly financial query so Pending/Paid amounts are reflected without
    # requiring the admin to visit the 13th Month page first.
    sync_thirteenth_month_records(year)
    sync_thirteenth_month_expenses(year)

    expenses = Expense.query.filter(
        Expense.expense_date >= start, Expense.expense_date < end
    ).all()
    paid_rows = [e for e in expenses if e.status == "Paid"]
    pending_rows = [e for e in expenses if e.status == "Pending"]

    # Commission upload is the company's income source. Total Collection is
    # income; Commission is an automatic company expense with no Paid/Pending
    # status in the Commission Report.
    commission_rows = CommissionRecord.query.all()
    selected_commissions = []
    for r in commission_rows:
        period = (r.period or "").strip().rstrip(".")
        matched = False
        for fmt in ("%B %Y", "%b %Y"):
            try:
                d = datetime.strptime(period, fmt)
                matched = (d.year == year and d.month == month)
                break
            except ValueError:
                pass
        if not matched:
            matched = period == f"{year}-{month:02d}"
        if matched:
            selected_commissions.append(r)

    commission_income = round(sum(float(r.total_collection or 0) for r in selected_commissions), 2)
    commission_expense = round(sum(float(r.commission or 0) for r in selected_commissions), 2)
    income_total = commission_income
    paid_expenses = round(sum(float(e.amount or 0) for e in paid_rows) + commission_expense, 2)
    pending_expenses = round(sum(float(e.amount or 0) for e in pending_rows), 2)
    profit = round(income_total - paid_expenses, 2)

    def cat_total(rows, category):
        return round(sum(float(e.amount or 0) for e in rows if e.category == category), 2)

    def gov_total(rows, label):
        return round(sum(
            float(e.amount or 0) for e in rows
            if e.source_type == "payroll_contribution" and label.lower() in (e.name or "").lower()
        ), 2)

    # Pending government liability is employee amount due + employer share.
    contrib_rows = (MonthlyContribution.query
                    .join(Payroll, MonthlyContribution.source_payroll_id == Payroll.id)
                    .filter(
                        MonthlyContribution.contribution_year == year,
                        MonthlyContribution.contribution_month == month,
                        Payroll.status.in_(["Approved", "Paid"]),
                        Payroll.period_start == date(year, month, 16),
                        Payroll.period_end == date(year, month, calendar.monthrange(year, month)[1]),
                    ).all())
    pending_gov = {}
    for label, field in (("sss", "sss"), ("philhealth", "philhealth"), ("pagibig", "pagibig")):
        gov_expense = Expense.query.filter_by(
            source_type="payroll_contribution",
            source_key=f"{year}-{month:02d}:{field}"
        ).first()
        if gov_expense is not None and gov_expense.status == "Paid":
            pending_gov[label] = 0.0
        else:
            pending_gov[label] = round(sum(
                float(getattr(r, f"{field}_due") or 0) +
                float(getattr(r, f"{field}_employer") or 0)
                for r in contrib_rows
            ), 2)
    pending_gov_total = round(sum(pending_gov.values()), 2)
    non_gov_pending = round(sum(
        float(e.amount or 0) for e in pending_rows if e.source_type != "payroll_contribution"
    ), 2)
    pending_expenses = round(non_gov_pending + pending_gov_total, 2)

    paid_breakdown = {
        "salary": cat_total(paid_rows, "Salary / Payroll"),
        "commission": commission_expense,
        "sss": gov_total(paid_rows, "sss"),
        "philhealth": gov_total(paid_rows, "philhealth"),
        "pagibig": gov_total(paid_rows, "pag-ibig"),
        "thirteenth_month": cat_total(paid_rows, "13th Month Pay"),
        "tax": cat_total(paid_rows, "Tax"),
        "other": round(sum(float(e.amount or 0) for e in paid_rows if e.category not in {"Salary / Payroll", "Tax", "Government", "13th Month Pay"} and e.source_type != "payroll_contribution"), 2),
        "government_total": round(sum(float(e.amount or 0) for e in paid_rows if e.source_type == "payroll_contribution"), 2),
    }
    pending_breakdown = {
        "salary": cat_total(pending_rows, "Salary / Payroll"),
        "sss": pending_gov["sss"],
        "philhealth": pending_gov["philhealth"],
        "pagibig": pending_gov["pagibig"],
        "thirteenth_month": cat_total(pending_rows, "13th Month Pay"),
        "tax": cat_total(pending_rows, "Tax"),
        "other": round(sum(float(e.amount or 0) for e in pending_rows if e.category not in {"Salary / Payroll", "Tax", "Government", "13th Month Pay"} and e.source_type != "payroll_contribution"), 2),
        "government_total": pending_gov_total,
    }

    return render_template(
        "financial_summary.html", year=year, month=month, incomes=[],
        income_total=income_total, paid_expenses=paid_expenses,
        pending_expenses=pending_expenses, profit=profit,
        paid_breakdown=paid_breakdown, pending_breakdown=pending_breakdown,
        commission_total=commission_expense, commission_income=commission_income
    )


@app.route("/income/delete/<int:income_id>", methods=["POST"])
@admin_required
def delete_income(income_id):
    income = FinancialIncome.query.get_or_404(income_id)
    year, month = income.income_date.year, income.income_date.month
    db.session.delete(income); db.session.commit(); flash("Income deleted.", "success")
    return redirect(url_for("financial_summary", year=year, month=month))

@app.route("/commission-report", methods=["GET", "POST"])
@admin_required
def commission_report():
    if request.method == "POST":
        period = request.form.get("period", "").strip()
        replace_existing = request.form.get("replace_existing") == "1"
        upload = request.files.get("commission_file")
        if not period:
            flash("Please enter the commission period, e.g. September 2026.", "danger")
            return redirect(url_for("commission_report"))
        if not upload or not upload.filename.lower().endswith((".xlsx", ".xlsm")):
            flash("Please upload an Excel file (.xlsx or .xlsm).", "danger")
            return redirect(url_for("commission_report"))

        try:
            from openpyxl import load_workbook
            upload.stream.seek(0)
            wb = load_workbook(upload.stream, data_only=True, read_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                raise ValueError("The Excel file is empty.")
            headers = [str(x).strip().lower() if x is not None else "" for x in rows[0]]
            required = ["agent name", "employee id", "total collection", "commission rate", "commission"]
            missing = [h for h in required if h not in headers]
            if missing:
                raise ValueError("Missing required column(s): " + ", ".join(missing))
            idx = {h: headers.index(h) for h in required}
            parsed = []
            errors = []
            for row_num, row in enumerate(rows[1:], start=2):
                if not any(v is not None and str(v).strip() for v in row):
                    continue
                try:
                    agent_name = str(row[idx["agent name"]] or "").strip()
                    employee_id = str(row[idx["employee id"]] or "").strip()
                    if not agent_name or not employee_id:
                        raise ValueError("Agent Name and Employee ID are required")
                    total = _money(row[idx["total collection"]])
                    rate = _parse_commission_rate(row[idx["commission rate"]])
                    commission = _money(row[idx["commission"]])
                    expected = _money(total * rate)
                    if abs(expected - commission) > 0.01:
                        raise ValueError(f"Commission {commission:,.2f} does not match {total:,.2f} × {rate:.2%} = {expected:,.2f}")
                    parsed.append((agent_name, employee_id, total, rate, commission))
                except Exception as exc:
                    errors.append(f"Row {row_num}: {exc}")
            if errors:
                raise ValueError("Upload rejected. " + " | ".join(errors[:10]) + (" | ..." if len(errors) > 10 else ""))
            if not parsed:
                raise ValueError("No commission records were found.")

            if replace_existing:
                CommissionRecord.query.filter_by(period=period).delete(synchronize_session=False)
            for agent_name, employee_id, total, rate, commission in parsed:
                db.session.add(CommissionRecord(
                    period=period,
                    agent_name=agent_name,
                    employee_id=employee_id,
                    total_collection=total,
                    commission_rate=rate,
                    commission=commission,
                ))
            db.session.commit()
            flash(f"Imported {len(parsed)} commission record(s) for {period}.", "success")
        except Exception as exc:
            db.session.rollback()
            flash(str(exc), "danger")
        return redirect(url_for("commission_report", period=period))

    period = request.args.get("period", "").strip()
    periods = [r[0] for r in db.session.query(CommissionRecord.period).distinct().order_by(CommissionRecord.period.desc()).all()]
    records = CommissionRecord.query.filter_by(period=period).order_by(CommissionRecord.agent_name.asc()).all() if period else []
    total_collection = sum(r.total_collection for r in records)
    total_commission = sum(r.commission for r in records)
    return render_template(
        "commission_report.html",
        records=records,
        periods=periods,
        selected_period=period,
        total_collection=total_collection,
        total_commission=total_commission,
    )


@app.route("/commission-template")
@admin_required
def commission_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Commission Upload"
    headers = ["Agent Name", "Employee ID", "Total Collection", "Commission Rate", "Commission"]
    ws.append(headers)
    examples = [
        ["Juan Dela Cruz", "EMP001", 500000, 0.10, 50000],
        ["Maria Santos", "EMP002", 380000, 0.05, 19000],
    ]
    for row in examples:
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for cell in ws["D"][1:]:
        cell.number_format = "0.00%"
    for cell in ws["C"][1:]:
        cell.number_format = '#,##0.00'
    for cell in ws["E"][1:]:
        cell.number_format = '#,##0.00'
    widths = {"A": 24, "B": 16, "C": 20, "D": 18, "E": 18}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name="commission_upload_template.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/commission-slip/<int:record_id>")
@admin_required
def commission_slip(record_id):
    record = CommissionRecord.query.get_or_404(record_id)
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, rightMargin=20*mm, leftMargin=20*mm, topMargin=20*mm, bottomMargin=20*mm)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("UNITY PATH RECOVERY AND COLLECTION SERVICES OPC", styles["Title"]),
        Spacer(1, 6),
        Paragraph("AGENT COMMISSION SLIP", styles["Heading2"]),
        Spacer(1, 12),
        Paragraph(f"<b>Commission Period:</b> {record.period}", styles["Normal"]),
        Paragraph(f"<b>Agent Name:</b> {record.agent_name}", styles["Normal"]),
        Paragraph(f"<b>Employee ID:</b> {record.employee_id}", styles["Normal"]),
        Spacer(1, 14),
    ]
    data = [
        ["Particular", "Amount"],
        ["Total Collection", f"₱{record.total_collection:,.2f}"],
        ["Commission Rate", f"{record.commission_rate:.2%}"],
        ["Commission Earned", f"₱{record.commission:,.2f}"],
    ]
    table = Table(data, colWidths=[110*mm, 55*mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#17365D")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("ALIGN", (1,1), (1,-1), "RIGHT"),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#EAF2F8")),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(table)
    story.append(Spacer(1, 24))
    story.append(Paragraph("This document is system-generated based on the uploaded commission report.", styles["Normal"]))
    doc.build(story)
    out.seek(0)
    filename = secure_filename(f"commission_slip_{record.employee_id}_{record.period}.pdf")
    return send_file(out, as_attachment=True, download_name=filename, mimetype="application/pdf")


@app.route("/commission-slips/<period>")
@admin_required
def commission_slips(period):
    records = CommissionRecord.query.filter_by(period=period).order_by(CommissionRecord.agent_name.asc()).all()
    if not records:
        flash("No commission records found for that period.", "danger")
        return redirect(url_for("commission_report"))
    bundle = BytesIO()
    with ZipFile(bundle, "w", ZIP_DEFLATED) as zf:
        for record in records:
            out = BytesIO()
            doc = SimpleDocTemplate(out, pagesize=A4, rightMargin=20*mm, leftMargin=20*mm, topMargin=20*mm, bottomMargin=20*mm)
            styles = getSampleStyleSheet()
            story = [Paragraph("UNITY PATH RECOVERY AND COLLECTION SERVICES OPC", styles["Title"]), Spacer(1, 6), Paragraph("AGENT COMMISSION SLIP", styles["Heading2"]), Spacer(1, 12), Paragraph(f"<b>Commission Period:</b> {record.period}", styles["Normal"]), Paragraph(f"<b>Agent Name:</b> {record.agent_name}", styles["Normal"]), Paragraph(f"<b>Employee ID:</b> {record.employee_id}", styles["Normal"]), Spacer(1, 14)]
            data = [["Particular", "Amount"], ["Total Collection", f"₱{record.total_collection:,.2f}"], ["Commission Rate", f"{record.commission_rate:.2%}"], ["Commission Earned", f"₱{record.commission:,.2f}"]]
            table = Table(data, colWidths=[110*mm, 55*mm])
            table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#17365D")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.5, colors.grey), ("ALIGN", (1,1), (1,-1), "RIGHT"), ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"), ("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#EAF2F8")), ("BOTTOMPADDING", (0,0), (-1,-1), 8), ("TOPPADDING", (0,0), (-1,-1), 8)]))
            story.append(table); story.append(Spacer(1, 24)); story.append(Paragraph("This document is system-generated based on the uploaded commission report.", styles["Normal"]))
            doc.build(story); out.seek(0)
            zf.writestr(secure_filename(f"commission_slip_{record.employee_id}.pdf"), out.getvalue())
    bundle.seek(0)
    return send_file(bundle, as_attachment=True, download_name=secure_filename(f"commission_slips_{period}.zip"), mimetype="application/zip")


@app.route("/dashboard")
@admin_required
def dashboard():
    total_employees = Employee.query.count()
    active_employees = Employee.query.filter_by(
        employment_status="Active"
    ).count()
    inactive_employees = Employee.query.filter_by(
        employment_status="Inactive"
    ).count()
    total_payroll = Payroll.query.count()
    sync_contribution_expenses_all_months()
    sync_payroll_salary_expenses()
    # Keep all current-year 13th Month records mirrored into expenses so the
    # overall dashboard Pending/Paid totals stay synchronized.
    sync_thirteenth_month_records(date.today().year)
    sync_thirteenth_month_expenses(date.today().year)
    income_total = round(sum(float(r.total_collection or 0) for r in CommissionRecord.query.all()), 2)
    commission_total = round(sum(float(r.commission or 0) for r in CommissionRecord.query.all()), 2)
    expenses = Expense.query.all()
    paid_expenses = round(sum(float(e.amount or 0) for e in expenses if e.status == "Paid") + commission_total, 2)
    pending_expenses = round(sum(float(e.amount or 0) for e in expenses if e.status == "Pending"), 2)
    profit = round(income_total - paid_expenses, 2)
    paid_payroll = round(sum(float(p.net_pay or 0) for p in Payroll.query.filter_by(status="Paid").all()), 2)

    return render_template(
        "dashboard.html",
        total_employees=total_employees, active_employees=active_employees, inactive_employees=inactive_employees,
        total_payroll=total_payroll, income_total=income_total, paid_expenses=paid_expenses,
        pending_expenses=pending_expenses, profit=profit, commission_total=commission_total, paid_payroll=paid_payroll
    )



REQUIRED_EMPLOYEE_FIELDS = [
    ("Employee ID", "employee_id"),
    ("First Name", "first_name"),
    ("Last Name", "last_name"),
    ("Department", "department"),
    ("Position / Job Title", "position"),
    ("Date Hired", "date_hired"),
    ("Employment Type", "employment_type"),
    ("Basic Daily Rate", "daily_rate"),
    ("Date of Birth", "date_of_birth"),
    ("Sex", "sex"),
    ("Civil Status", "civil_status"),
    ("Nationality", "nationality"),
    ("Current Home Address", "address"),
    ("Primary Mobile Number", "phone"),
    ("Personal Email", "email"),
    ("Emergency Contact Name", "emergency_name"),
    ("Emergency Contact Relationship", "emergency_relationship"),
    ("Emergency Contact Number", "emergency_phone"),
    ("SSS Number", "sss_number"),
    ("PhilHealth Number", "philhealth_number"),
    ("Pag-IBIG / MID Number", "pagibig_number"),
    ("TIN", "tin_number"),
    ("Bank Name", "bank_name"),
    ("Bank Account Name", "bank_account_name"),
    ("Bank Account Number", "bank_account_number"),
    ("2x2 Profile Photo", "profile_photo"),
]

REQUIRED_EMPLOYEE_DOCUMENTS = [
    "Resume / Curriculum Vitae",
    "Birth Certificate (PSA)",
    "NBI Clearance",
    "Police Clearance",
    "Barangay Clearance",
    "SSS ID / E-1 Form",
    "PhilHealth MDR Form",
    "Pag-IBIG MDF Form",
    "TIN / BIR Form 1902",
    "Valid Government ID",
    "2x2 ID Photos",
    "Medical Certificate",
    "Drug Test Result",
]


def employee_completion(employee):
    """Return a transparent profile/requirements checklist for the admin UI."""
    missing_details = []
    for label, attr in REQUIRED_EMPLOYEE_FIELDS:
        value = getattr(employee, attr, None)
        if value is None or (isinstance(value, str) and not value.strip()):
            missing_details.append(label)

    documents = _json_load(getattr(employee, "documents_json", None), [])
    submitted = {
        str(item.get("document", "")).strip()
        for item in documents
        if isinstance(item, dict) and str(item.get("submitted", "")).strip().lower() == "yes"
    }
    missing_documents = [
        name for name in REQUIRED_EMPLOYEE_DOCUMENTS
        if name not in submitted
    ]

    total = len(REQUIRED_EMPLOYEE_FIELDS) + len(REQUIRED_EMPLOYEE_DOCUMENTS)
    completed = total - len(missing_details) - len(missing_documents)
    percent = round((completed / total) * 100) if total else 100

    if not missing_details and not missing_documents:
        status = "Complete"
        status_class = "success"
    elif len(missing_details) + len(missing_documents) >= 8:
        status = "Critical Missing"
        status_class = "danger"
    else:
        status = "Incomplete"
        status_class = "warning"

    return {
        "percent": percent,
        "status": status,
        "status_class": status_class,
        "missing_details": missing_details,
        "missing_documents": missing_documents,
        "missing_count": len(missing_details) + len(missing_documents),
    }


@app.route("/employees")
@admin_required
def employees():
    records = Employee.query.order_by(Employee.last_name.asc()).all()
    completion = {employee.id: employee_completion(employee) for employee in records}
    return render_template("employees.html", employees=records, completion=completion)


def build_employee_profile_docx(emp):
    """Build a compact, professional 3-page branded A4 employee 201-file."""
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import RGBColor

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.27); sec.page_height = Inches(11.69)
    sec.top_margin = Inches(0.28); sec.bottom_margin = Inches(0.28)
    sec.left_margin = Inches(0.34); sec.right_margin = Inches(0.34)
    sec.header_distance = Inches(0.04); sec.footer_distance = Inches(0.05)
    NAVY, GOLD, LIGHT, DARK, WHITE, GRAY = "17365D", "C9A227", "EEF3F8", "263238", "FFFFFF", "6B7280"

    def val(v):
        return "" if v is None or str(v).strip().lower() in {"none", "null", "—", "-"} else str(v).strip()
    def money(v):
        try: return f"PHP {float(v or 0):,.2f}"
        except Exception: return ""
    def fdate(v): return v.strftime("%B %d, %Y") if v else ""
    def shade(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr(); shd = tcPr.find(qn('w:shd'))
        if shd is None: shd = OxmlElement('w:shd'); tcPr.append(shd)
        shd.set(qn('w:fill'), fill)
    def margins(cell, top=20, start=55, bottom=20, end=55):
        tcPr = cell._tc.get_or_add_tcPr(); mar = tcPr.first_child_found_in('w:tcMar')
        if mar is None: mar = OxmlElement('w:tcMar'); tcPr.append(mar)
        for m,v in [('top',top),('start',start),('bottom',bottom),('end',end)]:
            n = mar.find(qn('w:'+m))
            if n is None: n = OxmlElement('w:'+m); mar.append(n)
            n.set(qn('w:w'), str(v)); n.set(qn('w:type'), 'dxa')
    def width(cell, inches):
        cell.width = Inches(inches); tcPr = cell._tc.get_or_add_tcPr(); tcW = tcPr.find(qn('w:tcW'))
        if tcW is None: tcW = OxmlElement('w:tcW'); tcPr.append(tcW)
        tcW.set(qn('w:w'), str(int(inches*1440))); tcW.set(qn('w:type'), 'dxa')
    def borderless(table):
        tblPr = table._tbl.tblPr; b = tblPr.first_child_found_in('w:tblBorders')
        if b is None: b = OxmlElement('w:tblBorders'); tblPr.append(b)
        for edge in ('top','left','bottom','right','insideH','insideV'):
            n = b.find(qn('w:'+edge))
            if n is None: n = OxmlElement('w:'+edge); b.append(n)
            n.set(qn('w:val'), 'nil')
    def borders(table, color="CBD5E1", size="4"):
        tblPr = table._tbl.tblPr; b = tblPr.first_child_found_in('w:tblBorders')
        if b is None: b = OxmlElement('w:tblBorders'); tblPr.append(b)
        for edge in ('top','left','bottom','right','insideH','insideV'):
            n = b.find(qn('w:'+edge))
            if n is None: n = OxmlElement('w:'+edge); b.append(n)
            n.set(qn('w:val'),'single'); n.set(qn('w:sz'),size); n.set(qn('w:color'),color)
    def font(run, size=7.4, bold=False, color=DARK):
        run.font.name='Aptos'; run.font.size=Pt(size); run.bold=bold; run.font.color.rgb=RGBColor.from_string(color)
    doc.styles['Normal'].font.name='Aptos'; doc.styles['Normal'].font.size=Pt(7.4)
    doc.styles['Normal'].paragraph_format.space_before=Pt(0); doc.styles['Normal'].paragraph_format.space_after=Pt(0); doc.styles['Normal'].paragraph_format.line_spacing=1.0

    # Compact branded header/footer.
    h=sec.header; hp=h.paragraphs[0]; hp.alignment=WD_ALIGN_PARAGRAPH.CENTER; hp.paragraph_format.space_before=Pt(0); hp.paragraph_format.space_after=Pt(0)
    logo_path=os.path.join(BASE_DIR,'static','logo.png')
    if os.path.exists(logo_path):
        rr=hp.add_run(); rr.add_picture(logo_path, width=Inches(1.05))
    p=h.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
    r=p.add_run('UNITY PATH RECOVERY AND COLLECTION SERVICES OPC'); font(r,8.4,True,NAVY)
    p=h.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(1)
    r=p.add_run('EMPLOYEE 201 FILE  •  CONFIDENTIAL  •  FOR INTERNAL HR USE ONLY'); font(r,6.4,True,GRAY)
    p=h.add_paragraph(); p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
    pPr=p._p.get_or_add_pPr(); pbdr=OxmlElement('w:pBdr'); bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'12'); bot.set(qn('w:space'),'0'); bot.set(qn('w:color'),GOLD); pbdr.append(bot); pPr.append(pbdr)
    f=sec.footer.paragraphs[0]; f.alignment=WD_ALIGN_PARAGRAPH.CENTER; f.paragraph_format.space_before=Pt(0); f.paragraph_format.space_after=Pt(0); r=f.add_run(f'UPRCSO  •  {val(emp.employee_id)}  •  Confidential'); font(r,6.2,False,GRAY)

    education=_json_load(emp.education_json,[]); work=_json_load(emp.work_experience_json,[]); skills=_json_load(emp.skills_json,{}); family=_json_load(emp.family_json,[]); refs=_json_load(emp.references_json,[]); docs=_json_load(emp.documents_json,[])
    full=' '.join(x for x in [val(emp.first_name),val(emp.middle_name),val(emp.last_name)] if x)

    def compact_spacing(paragraph, after=0, before=0, line=1.0):
        paragraph.paragraph_format.space_before=Pt(before); paragraph.paragraph_format.space_after=Pt(after); paragraph.paragraph_format.line_spacing=line

    def section(title, num=None):
        t=doc.add_table(rows=1,cols=1); t.autofit=False; borders(t,NAVY,'4')
        c=t.cell(0,0); shade(c,NAVY); margins(c,25,65,25,65)
        p=c.paragraphs[0]; compact_spacing(p); r=p.add_run((f'{num}  ' if num else '')+title.upper()); font(r,7.5,True,WHITE)
        p=doc.add_paragraph(); compact_spacing(p,after=1)

    def fields(rows, widths=(2.05,5.0)):
        t=doc.add_table(rows=0,cols=2); t.autofit=False; borders(t)
        for label,value in rows:
            c=t.add_row().cells; width(c[0],widths[0]); width(c[1],widths[1]); shade(c[0],LIGHT)
            margins(c[0]); margins(c[1]); c[0].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; c[1].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for cell in c:
                p=cell.paragraphs[0]; p.clear(); compact_spacing(p)
            r=c[0].paragraphs[0].add_run(label); font(r,6.8,True,NAVY)
            r=c[1].paragraphs[0].add_run(val(value) or '—'); font(r,6.9)
        return t

    def grid(headers, rows, widths, body_size=6.35):
        t=doc.add_table(rows=1,cols=len(headers)); t.autofit=False; borders(t); t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
        for i,hdr in enumerate(headers):
            c=t.cell(0,i); width(c,widths[i]); shade(c,NAVY); margins(c,25,35,25,35); p=c.paragraphs[0]; p.clear(); compact_spacing(p); r=p.add_run(hdr); font(r,6.3,True,WHITE)
        for vals in rows:
            cells=t.add_row().cells
            for i,v in enumerate(vals):
                width(cells[i],widths[i]); margins(cells[i],18,30,18,30); p=cells[i].paragraphs[0]; p.clear(); compact_spacing(p); r=p.add_run(val(v) or '—'); font(r,body_size, i==0, NAVY if i==0 else DARK); shade(cells[i],LIGHT if i==0 else WHITE)
        return t

    def spacer(h=1):
        p=doc.add_paragraph(); compact_spacing(p,after=h)

    def page_break():
        p=doc.add_paragraph(); compact_spacing(p); p.add_run().add_break()
        # Explicit page break keeps the document exactly 3 logical pages.
        p.runs[0].add_break()

    # PAGE 1 — identity, employment, personal, contact.
    section('Employee Identity')
    t=doc.add_table(rows=1,cols=2); t.autofit=False; borderless(t); width(t.cell(0,0),1.25); width(t.cell(0,1),5.8)
    c=t.cell(0,0); margins(c,20,20,20,20)
    if emp.profile_photo:
        path=os.path.join(PROFILE_DIR,emp.profile_photo)
        if os.path.exists(path): c.paragraphs[0].add_run().add_picture(path,width=Inches(1.02),height=Inches(1.02))
    c=t.cell(0,1); margins(c,15,80,15,40); p=c.paragraphs[0]; compact_spacing(p,after=1); r=p.add_run(full.upper()); font(r,15,True,NAVY)
    p=c.add_paragraph(); compact_spacing(p,after=1); r=p.add_run(val(emp.position) or 'EMPLOYEE'); font(r,8.5,True,GOLD)
    p=c.add_paragraph(); compact_spacing(p,after=1); r=p.add_run(f'Employee ID: {val(emp.employee_id)}   •   Status: {val(emp.employment_status) or "Active"}'); font(r,7.4,True,DARK)
    p=c.add_paragraph(); compact_spacing(p); r=p.add_run(val(emp.professional_summary) or 'Professional employee profile and personnel record.'); font(r,6.8,False,GRAY)
    spacer(1); section('Employment Information','1')
    fields([('Employee ID',emp.employee_id),('Date Hired',fdate(emp.date_hired)),('Employment Status',emp.employment_status),('Department',emp.department),('Position / Job Title',emp.position),('Reporting To',emp.reporting_to),('Employment Type',emp.employment_type),('Work Location',emp.work_location),('Work Schedule',emp.schedule_type),('Schedule Details',emp.schedule_details or 'Monday–Friday, 8:00 AM–5:00 PM')])
    spacer(1); section('Personal Information','2')
    # Two-column personal matrix saves vertical space.
    personal_rows=[
        [('First Name',emp.first_name),('Middle Name',emp.middle_name),('Last Name',emp.last_name),('Nickname / Preferred Name',emp.nickname)],
        [('Date of Birth',emp.date_of_birth),('Place of Birth',emp.place_of_birth),('Age',getattr(emp,'age','')),('Sex',emp.sex)],
        [('Civil Status',emp.civil_status),('Nationality',emp.nationality),('Religion',emp.religion),('Blood Type',emp.blood_type)],
        [('Height',emp.height),('Weight',emp.weight),('',''),('','')],
    ]
    t=doc.add_table(rows=0,cols=4); t.autofit=False; borders(t)
    for row in personal_rows:
        cells=t.add_row().cells
        for i,(label,value) in enumerate(row):
            width(cells[i],1.76); margins(cells[i],18,30,18,30); p=cells[i].paragraphs[0]; p.clear(); compact_spacing(p)
            if label:
                shade(cells[i],LIGHT); r=p.add_run(label+': '); font(r,6.4,True,NAVY); r=p.add_run(val(value) or '—'); font(r,6.4)
            else: shade(cells[i],WHITE)
    spacer(1); section('Contact Information','3')
    fields([('Current Home Address',emp.address),('Permanent Address',emp.permanent_address),('Mobile Number',emp.phone),('Secondary Number',emp.secondary_phone),('Personal Email',emp.email),('Company Email',emp.company_email),('Emergency Contact',emp.emergency_name),('Emergency Number',emp.emergency_phone),('Relationship',emp.emergency_relationship),('Emergency Address',emp.emergency_address)])

    # PAGE 2 — government, education, work, skills.
    page_break(); section('Government IDs & Numbers','4')
    gov=[('SSS Number',emp.sss_number),('PhilHealth Number',emp.philhealth_number),('Pag-IBIG / MID Number',emp.pagibig_number),('TIN',emp.tin_number),('UMID Number',emp.umid_number),('PhilSys / National ID',emp.national_id),("Voter's ID",emp.voters_id),("Driver's License",emp.drivers_license),('Passport Number',emp.passport_number),('Passport Expiry',emp.passport_expiry)]
    # Government IDs in a 4-column matrix.
    t=doc.add_table(rows=0,cols=4); t.autofit=False; borders(t)
    for i in range(0,len(gov),2):
        cells=t.add_row().cells
        pairs=gov[i:i+2]+[('', '')]*(2-len(gov[i:i+2]))
        for j,(label,value) in enumerate(pairs):
            for k,text in ((j*2,label),(j*2+1,val(value) or '—')):
                width(cells[k],1.76); margins(cells[k],18,30,18,30); p=cells[k].paragraphs[0]; p.clear(); compact_spacing(p); shade(cells[k],LIGHT if k%2==0 else WHITE); r=p.add_run(text); font(r,6.45,k%2==0,NAVY if k%2==0 else DARK)
    spacer(1); section('Educational Background','5')
    erows=[[x.get('level'),x.get('school'),x.get('course'),x.get('year'),x.get('honors')] for x in education]
    if not erows: erows=[['Elementary','','','',''],['Junior High School','','','',''],['Senior High School','','','',''],['Vocational / TESDA','','','',''],['College / University','','','',''],['Post-Graduate','','','','']]
    grid(['Level','School / Institution','Degree / Course','Year','Honors / Awards'],erows,[1.15,2.25,1.45,.72,1.48],6.15)
    spacer(1); section('Work Experience','6')
    wrows=[[x.get('company'),x.get('position'),x.get('from'),x.get('to'),x.get('reason')] for x in work]
    if not wrows: wrows=[['','','','',''] for _ in range(3)]
    grid(['Company / Employer','Position / Title','Date From','Date To','Reason for Leaving'],wrows,[1.72,1.42,1.0,1.0,1.91],6.15)
    spacer(1); section('Skills & Competencies','7')
    fields([('Collection / Recovery Skills',skills.get('collection')),('Negotiation Skills',skills.get('negotiation')),('Computer Skills',skills.get('computer')),('Communication — Oral',skills.get('oral')),('Communication — Written',skills.get('written')),('Languages Spoken',skills.get('languages')),('Other Skills / Certifications',skills.get('other')),('TESDA / PRC Licenses',skills.get('licenses'))])

    # PAGE 3 — job, compensation, family, references, checklist, declaration.
    page_break(); section('Job Description','8')
    fields([('Position Title',emp.position),('Department',emp.department),('Reports To',emp.reporting_to),('Employment Type',emp.employment_type),('Work Location',emp.work_location),('Work Schedule',emp.schedule_type),('Schedule Details',emp.schedule_details)])
    p=doc.add_paragraph(); compact_spacing(p,before=2,after=1); r=p.add_run('Primary Duties & Responsibilities'); font(r,7.3,True,NAVY)
    duty_lines=[x.strip() for x in (emp.duties or '').splitlines() if x.strip()]
    if duty_lines:
        for i,x in enumerate(duty_lines[:6],1):
            p=doc.add_paragraph(f'{i}. {x}'); compact_spacing(p,after=1); font(p.runs[0],6.8)
    else:
        for i in range(1,4):
            p=doc.add_paragraph(f'{i}. ' + '_'*90); compact_spacing(p,after=0); font(p.runs[0],6.5,False,GRAY)
    spacer(1); section('Compensation & Benefits','9')
    fields([('Basic Daily Rate',money(emp.daily_rate)),('Daily Allowance',money(emp.daily_allowance)),('Daily Incentive',money(emp.daily_incentive)),('Salary Frequency','Daily / Attendance-Based')])
    spacer(1); section('Family Background','10')
    frows=[[x.get('relationship'),x.get('name'),x.get('dob'),x.get('occupation'),x.get('contact')] for x in family]
    if not frows: frows=[['Father','','','',''],['Mother','','','',''],['Spouse','','','',''],['Child 1','','','',''],['Child 2','','','','']]
    grid(['Relationship','Full Name','Date of Birth','Occupation','Contact Number'],frows,[1.0,1.9,1.15,1.4,1.6],6.0)
    spacer(1); section('Character References','11')
    rrows=[[x.get('name'),x.get('company'),x.get('position'),x.get('contact')] for x in refs]
    if not rrows: rrows=[['','','',''] for _ in range(2)]
    grid(['Full Name','Company / Organization','Position','Contact Number'],rrows,[1.65,2.2,1.35,1.85],6.0)
    spacer(1); section('Pre-Employment Documents Checklist','12')
    default_docs=['Resume / Curriculum Vitae','Birth Certificate (PSA)','NBI Clearance','Police Clearance','Barangay Clearance','SSS ID / E-1 Form','PhilHealth MDR Form','Pag-IBIG MDF Form','TIN / BIR Form 1902','Valid Government ID','2x2 ID Photos','Certificate of Employment (previous)','Diploma / Transcript of Records','Medical Certificate','Drug Test Result']
    dmap={x.get('document'):x for x in docs}; names=list(dict.fromkeys(default_docs+[x.get('document') for x in docs if x.get('document')]))
    # Two-column checklist for a compact one-block section.
    t=doc.add_table(rows=0,cols=4); t.autofit=False; borders(t)
    left=names[:(len(names)+1)//2]; right=names[(len(names)+1)//2:]
    for i in range(max(len(left),len(right))):
        cells=t.add_row().cells
        for side,name in enumerate([left[i] if i<len(left) else '', right[i] if i<len(right) else '']):
            a=side*2; b=a+1
            width(cells[a],2.65); width(cells[b],.88); margins(cells[a],15,25,15,25); margins(cells[b],15,25,15,25)
            shade(cells[a],LIGHT); shade(cells[b],WHITE)
            p=cells[a].paragraphs[0]; p.clear(); compact_spacing(p); r=p.add_run(('☑ ' if name and dmap.get(name,{}).get('submitted')=='Yes' else '☐ ')+name); font(r,5.9,False,DARK)
            p=cells[b].paragraphs[0]; p.clear(); compact_spacing(p); r=p.add_run('Yes' if name and dmap.get(name,{}).get('submitted')=='Yes' else 'No'); font(r,5.8,True,NAVY)
    spacer(1); section('Employee Declaration','13')
    p=doc.add_paragraph('I hereby certify that all information provided in this Employee Profile is true, correct, and complete to the best of my knowledge. I understand that any false or misleading information may result in appropriate employment action.'); compact_spacing(p,after=7); font(p.runs[0],6.6,False,DARK)
    sig=doc.add_table(rows=2,cols=2); sig.autofit=False; borderless(sig)
    sig.cell(0,0).text='____________________________\nEmployee Signature'; sig.cell(0,1).text='____________________________\nReceived & Verified by HR'; sig.cell(1,0).text='Date: __________________'; sig.cell(1,1).text='Date: __________________'
    for row in sig.rows:
        for c in row.cells:
            margins(c,20,45,20,45)
            for pp in c.paragraphs: compact_spacing(pp)
            for rr in c.paragraphs[0].runs: font(rr,6.4,'Signature' in c.text or 'HR' in c.text,NAVY if ('Signature' in c.text or 'HR' in c.text) else DARK)
    return doc

@app.route("/employees/export-docx", methods=["GET", "POST"])
@admin_required
def export_employees_docx():
    """Export selected employees as separate DOCX files; multiple files are bundled into one ZIP."""
    if request.method == "POST":
        selected = request.form.getlist("selected_employee_ids")
    else:
        selected = request.args.getlist("selected_employee_ids")

    ids=[]
    for value in selected:
        try: ids.append(int(value))
        except (TypeError, ValueError): pass

    if not ids:
        flash("Please select at least one employee to export.", "warning")
        return redirect(url_for("employees"))

    employees = Employee.query.filter(Employee.id.in_(ids)).order_by(Employee.last_name.asc(), Employee.first_name.asc()).all()
    if not employees:
        flash("No selected employees were found.", "warning")
        return redirect(url_for("employees"))

    files=[]
    for emp in employees:
        doc=build_employee_profile_docx(emp)
        out=BytesIO(); doc.save(out); out.seek(0)
        safe=secure_filename(f"{emp.employee_id}_{emp.first_name}_{emp.last_name}_Employee_Profile.docx")
        files.append((safe, out.getvalue()))

    if len(files) == 1:
        filename=files[0][0]
        return send_file(BytesIO(files[0][1]), as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    archive=BytesIO()
    with ZipFile(archive, "w", ZIP_DEFLATED) as z:
        for filename,data in files: z.writestr(filename,data)
    archive.seek(0)
    return send_file(archive, as_attachment=True, download_name=f"employee_profiles_{date.today().isoformat()}.zip", mimetype="application/zip")


@app.route("/employees/add", methods=["GET", "POST"])
@admin_required
def add_employee():
    employee = Employee()

    if request.method == "POST":
        try:
            save_employee_form(employee)
            if not employee.employee_id or not employee.first_name or not employee.last_name:
                raise ValueError("Employee ID, first name, and last name are required.")

            db.session.add(employee)
            db.session.commit()

            photo = request.files.get("profile_photo")
            if photo:
                save_employee_photo = photo
                save_profile_photo(employee, save_employee_photo)
                db.session.commit()

            flash("Employee added successfully.", "success")
            return redirect(url_for("employee_profile", id=employee.id))

        except Exception as exc:
            db.session.rollback()
            flash(str(exc), "danger")

    return render_template("employee_form.html", employee=None, profile_data={"education": [], "work": [], "skills": {}, "family": [], "references": [], "documents": []})


@app.route("/employees/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_employee(id):
    employee = Employee.query.get_or_404(id)

    if request.method == "POST":
        try:
            save_employee_form(employee)
            photo = request.files.get("profile_photo")
            if photo:
                save_profile_photo(employee, photo)

            db.session.commit()
            flash("Employee updated successfully.", "success")
            return redirect(url_for("employee_profile", id=employee.id))

        except Exception as exc:
            db.session.rollback()
            flash(str(exc), "danger")

    return render_template("employee_form.html", employee=employee, profile_data={
        "education": _json_load(employee.education_json, []),
        "work": _json_load(employee.work_experience_json, []),
        "skills": _json_load(employee.skills_json, {}),
        "family": _json_load(employee.family_json, []),
        "references": _json_load(employee.references_json, []),
        "documents": _json_load(employee.documents_json, []),
    })


@app.route("/employees/<int:id>")
@admin_required
def employee_profile(id):
    employee = Employee.query.get_or_404(id)
    attendance = Attendance.query.filter_by(
        employee_id=employee.id
    ).order_by(Attendance.attendance_date.desc()).limit(20).all()

    return render_template(
        "employee_profile.html",
        employee=employee,
        attendance=attendance,
        completion=employee_completion(employee)
    )


@app.route("/employees/delete/<int:id>", methods=["POST"])
@admin_required
def delete_employee(id):
    employee = Employee.query.get_or_404(id)
    db.session.delete(employee)
    db.session.commit()
    flash("Employee deleted.", "success")
    return redirect(url_for("employees"))


@app.route("/employees/<int:id>/create-account", methods=["GET", "POST"])
@admin_required
def create_employee_account(id):
    employee = Employee.query.get_or_404(id)

    if employee.user_account:
        flash("Employee account already exists.", "warning")
        return redirect(url_for("employee_account", id=id))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        if not username or not password:
            flash("Username and password are required.", "danger")
            return render_template("create_employee_account.html", employee=employee)

        if len(password) < 8:
            flash("Password must be at least 8 characters.", "danger")
            return render_template("create_employee_account.html", employee=employee)

        if User.query.filter_by(username=username).first():
            flash("Username already exists.", "danger")
            return render_template("create_employee_account.html", employee=employee)

        user = User(
            username=username,
            role="employee",
            employee_id=employee.id
        )
        user.set_password(password)

        db.session.add(user)
        db.session.commit()

        flash("Employee account created.", "success")
        return redirect(url_for("employee_account", id=id))

    return render_template("create_employee_account.html", employee=employee)


@app.route("/employees/<int:id>/account", methods=["GET", "POST"])
@admin_required
def employee_account(id):
    """Admin-only account management. Existing passwords are never displayed."""
    employee = Employee.query.get_or_404(id)
    user = employee.user_account

    if not user:
        return redirect(url_for("create_employee_account", id=id))

    if request.method == "POST":
        action = request.form.get("action", "").strip()

        if action == "change_username":
            username = request.form.get("username", "").strip()
            if not username:
                flash("Username is required.", "danger")
            else:
                existing = User.query.filter(User.username == username, User.id != user.id).first()
                if existing:
                    flash("Username already exists.", "danger")
                else:
                    user.username = username
                    db.session.commit()
                    flash("Username updated successfully.", "success")
            return redirect(url_for("employee_account", id=id))

        if action == "reset_password":
            password = request.form.get("password", "")
            confirm = request.form.get("confirm_password", "")
            if len(password) < 8:
                flash("Password must be at least 8 characters.", "danger")
            elif password != confirm:
                flash("Password and confirmation do not match.", "danger")
            else:
                user.set_password(password)
                db.session.commit()
                flash("Password reset successfully. Give the new password to the employee securely.", "success")
            return redirect(url_for("employee_account", id=id))

        flash("Invalid account action.", "danger")
        return redirect(url_for("employee_account", id=id))

    return render_template(
        "employee_account.html",
        employee=employee,
        user=user,
        completion=employee_completion(employee)
    )


@app.route("/calendar")
@admin_required
def calendar_view():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if month < 1 or month > 12:
        month = today.month
    holidays = Holiday.query.filter(
        Holiday.holiday_date >= date(year, month, 1),
        Holiday.holiday_date <= date(year, month, calendar.monthrange(year, month)[1])
    ).order_by(Holiday.holiday_date.asc()).all()
    return render_template("calendar.html", year=year, month=month,
                           month_name=calendar.month_name[month], days_in_month=calendar.monthrange(year, month)[1], holidays=holidays,
                           prev_year=(year-1 if month == 1 else year),
                           prev_month=(12 if month == 1 else month-1),
                           next_year=(year+1 if month == 12 else year),
                           next_month=(1 if month == 12 else month+1),
                           holiday_types=["Regular Holiday", "Special Non-Working Day", "Special Working Day", "Company Holiday", "Rest Day"])


@app.route("/calendar/add", methods=["POST"])
@admin_required
def add_holiday():
    d = parse_date(request.form.get("holiday_date"))
    name = request.form.get("name", "").strip()
    htype = request.form.get("holiday_type", "Regular Holiday")
    if not d or not name:
        flash("Holiday date and name are required.", "danger")
        return redirect(url_for("calendar_view"))
    if Holiday.query.filter_by(holiday_date=d).first():
        flash("A holiday already exists on that date. Edit the existing entry instead.", "danger")
        return redirect(url_for("calendar_view", year=d.year, month=d.month))
    defaults = holiday_defaults(htype)
    db.session.add(Holiday(holiday_date=d, name=name, holiday_type=htype,
                           work_multiplier=float(request.form.get("work_multiplier") or defaults[0]),
                           unworked_multiplier=float(request.form.get("unworked_multiplier") or defaults[1]),
                           ot_multiplier=float(request.form.get("ot_multiplier") or defaults[2]), active=True))
    db.session.commit()
    flash("Holiday added to Calendar. Payroll will use this entry when generated/recalculated.", "success")
    return redirect(url_for("calendar_view", year=d.year, month=d.month))


@app.route("/calendar/<int:id>/edit", methods=["POST"])
@admin_required
def edit_holiday(id):
    h = Holiday.query.get_or_404(id)
    d = parse_date(request.form.get("holiday_date")) or h.holiday_date
    h.holiday_date = d
    h.name = request.form.get("name", "").strip() or h.name
    h.holiday_type = request.form.get("holiday_type", h.holiday_type)
    h.work_multiplier = max(0.0, float(request.form.get("work_multiplier") or 0))
    h.unworked_multiplier = max(0.0, float(request.form.get("unworked_multiplier") or 0))
    h.ot_multiplier = max(0.0, float(request.form.get("ot_multiplier") or 0))
    h.active = request.form.get("active") == "1"
    db.session.commit()
    flash("Holiday updated. Payroll will use the updated Calendar entry.", "success")
    return redirect(url_for("calendar_view", year=d.year, month=d.month))


@app.route("/calendar/<int:id>/delete", methods=["POST"])
@admin_required
def delete_holiday(id):
    h = Holiday.query.get_or_404(id)
    d = h.holiday_date
    db.session.delete(h)
    db.session.commit()
    flash("Holiday removed from Calendar.", "success")
    return redirect(url_for("calendar_view", year=d.year, month=d.month))


@app.route("/attendance")
@admin_required
def attendance():
    """Attendance viewer with whole-month and custom date-range filters.

    Query options:
      - month=YYYY-MM : displays the entire calendar month
      - start_date=YYYY-MM-DD&end_date=YYYY-MM-DD : displays a custom range
      - date=YYYY-MM-DD : backward-compatible single-day view
    """
    today = date.today()

    month_value = (request.args.get("month") or "").strip()
    start_date = parse_date(request.args.get("start_date"))
    end_date = parse_date(request.args.get("end_date"))
    legacy_date = parse_date(request.args.get("date"))

    if start_date and not end_date:
        end_date = start_date
    if end_date and not start_date:
        start_date = end_date

    # Month filter takes priority when supplied.
    if month_value:
        try:
            year, month = [int(x) for x in month_value.split("-", 1)]
            start_date = date(year, month, 1)
            end_date = date(year, month, calendar.monthrange(year, month)[1])
        except (ValueError, TypeError):
            month_value = ""

    # Preserve the old single-day URL format.
    if not month_value and not (start_date or end_date) and legacy_date:
        start_date = end_date = legacy_date

    # Default: whole current month.
    if not start_date or not end_date:
        start_date = date(today.year, today.month, 1)
        end_date = date(today.year, today.month, calendar.monthrange(today.year, today.month)[1])

    if start_date > end_date:
        start_date, end_date = end_date, start_date

    employees = Employee.query.filter_by(
        employment_status="Active"
    ).order_by(Employee.last_name.asc()).all()

    records = Attendance.query.filter(
        Attendance.attendance_date >= start_date,
        Attendance.attendance_date <= end_date
    ).order_by(Attendance.attendance_date.asc(), Attendance.employee_id.asc()).all()

    # One record per employee/date for the viewer. If duplicate rows exist,
    # keep the first completed record; otherwise keep the first available one.
    record_map = {}
    for record in records:
        key = (record.attendance_date, record.employee_id)
        current = record_map.get(key)
        if current is None:
            record_map[key] = record
        elif not current.time_out and record.time_out:
            record_map[key] = record

    dates = []
    cursor = start_date
    while cursor <= end_date:
        dates.append(cursor)
        cursor += timedelta(days=1)

    # Build rows for every active employee/date so missing Time In/Out is visible
    # as an absence rather than silently disappearing from the month view.
    attendance_rows = []
    for attendance_date in dates:
        for employee in employees:
            attendance_rows.append({
                "attendance_date": attendance_date,
                "employee": employee,
                "record": record_map.get((attendance_date, employee.id))
            })

    if start_date == end_date:
        range_label = start_date.strftime("%B %d, %Y")
    else:
        range_label = f"{start_date.strftime('%B %d, %Y')} – {end_date.strftime('%B %d, %Y')}"

    return render_template(
        "attendance.html",
        selected_date=start_date,
        start_date=start_date,
        end_date=end_date,
        month_value=(month_value or start_date.strftime("%Y-%m")),
        employees=employees,
        record_map=record_map,
        attendance_rows=attendance_rows,
        range_label=range_label
    )


@app.route("/attendance/export-excel")
@admin_required
def export_attendance_excel():
    """Export the currently selected attendance date range to an Excel workbook.

    The export mirrors the Attendance page: it includes every active employee
    for every date in the selected range, including dates with no attendance
    record (shown as Absent on the page).
    """
    today = date.today()
    month_value = (request.args.get("month") or "").strip()
    start_date = parse_date(request.args.get("start_date"))
    end_date = parse_date(request.args.get("end_date"))

    if start_date and not end_date:
        end_date = start_date
    if end_date and not start_date:
        start_date = end_date

    if month_value:
        try:
            year, month = [int(x) for x in month_value.split("-", 1)]
            start_date = date(year, month, 1)
            end_date = date(year, month, calendar.monthrange(year, month)[1])
        except (ValueError, TypeError):
            month_value = ""

    if not start_date or not end_date:
        start_date = date(today.year, today.month, 1)
        end_date = date(today.year, today.month, calendar.monthrange(today.year, today.month)[1])

    if start_date > end_date:
        start_date, end_date = end_date, start_date

    employees = Employee.query.filter_by(
        employment_status="Active"
    ).order_by(Employee.last_name.asc(), Employee.first_name.asc()).all()

    records = Attendance.query.filter(
        Attendance.attendance_date >= start_date,
        Attendance.attendance_date <= end_date
    ).order_by(Attendance.attendance_date.asc(), Attendance.employee_id.asc()).all()

    record_map = {}
    for record in records:
        key = (record.attendance_date, record.employee_id)
        current = record_map.get(key)
        if current is None:
            record_map[key] = record
        elif not current.time_out and record.time_out:
            record_map[key] = record

    dates = []
    cursor = start_date
    while cursor <= end_date:
        dates.append(cursor)
        cursor += timedelta(days=1)

    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False

    title = f"ATTENDANCE REPORT — {start_date.strftime('%B %d, %Y')} to {end_date.strftime('%B %d, %Y')}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
    ws.cell(1, 1, title)
    ws.cell(1, 1).font = Font(bold=True, size=14)
    ws.cell(1, 1).alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)
    ws.cell(2, 1, f"Generated: {datetime.now().strftime('%B %d, %Y %I:%M %p')}")
    ws.cell(2, 1).font = Font(italic=True, size=10)
    ws.cell(2, 1).alignment = Alignment(horizontal="center")

    headers = [
        "Date", "Employee ID", "Employee Name", "Department", "Position",
        "Time In", "Time Out", "Hours", "Late Hours", "Undertime Hours",
        "OT Available", "Approved OT", "Status"
    ]
    # 13 columns, so expand the merged title rows to cover all columns.
    ws.unmerge_cells("A1:L1")
    ws.unmerge_cells("A2:L2")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))

    header_row = 4
    thin = Side(style="thin", color="D9E1F2")
    for col, header in enumerate(headers, 1):
        cell = ws.cell(header_row, col, header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.border = Border(bottom=thin)

    row_num = 5
    for attendance_date in dates:
        for employee in employees:
            record = record_map.get((attendance_date, employee.id))
            if not record:
                status = "Absent"
            elif (
                record.time_out
                and record.time_in
                and (record.time_out.date() != record.attendance_date
                     or record.time_out.date() != record.time_in.date())
            ):
                status = "Invalid Out"
            elif record.time_out and not record.time_in:
                status = "Invalid Out"
            elif (record.status or "").upper() == "VL":
                status = "Vacation Leave (VL)"
            elif (record.status or "").upper() == "SL":
                status = "Sick Leave (SL)"
            elif not record.time_in:
                status = "Absent"
            elif not record.time_out:
                status = "Incomplete"
            else:
                status = "Present"

            values = [
                attendance_date,
                employee.employee_id or "",
                f"{employee.first_name or ''} {employee.last_name or ''}".strip(),
                employee.department or "",
                employee.position or "",
                record.time_in if record and record.time_in else None,
                record.time_out if record and record.time_out else None,
                float(record.total_hours or 0) if record and record.status == "Present" else 0.0,
                float(record.late_hours or 0) if record and record.status == "Present" else 0.0,
                float(record.undertime_hours or 0) if record and record.status == "Present" else 0.0,
                float(record.overtime_hours or 0) if record and record.status == "Present" else 0.0,
                float(record.approved_overtime_hours or 0) if record and record.status == "Present" else 0.0,
                status,
            ]
            for col, value in enumerate(values, 1):
                cell = ws.cell(row_num, col, value)
                cell.border = Border(bottom=thin)
                if col == 1 and value:
                    cell.number_format = "mmm d, yyyy"
                elif col in (6, 7) and value:
                    cell.number_format = "h:mm AM/PM"
                elif col in (8, 9, 10, 11, 12):
                    cell.number_format = "0.00"
            row_num += 1

    # Add an auto-filter and practical column widths.
    last_row = max(row_num - 1, header_row)
    ws.auto_filter.ref = f"A{header_row}:M{last_row}"
    widths = [14, 15, 24, 18, 20, 14, 14, 10, 12, 16, 14, 14, 13]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[4].height = 24

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"attendance_{start_date.isoformat()}_to_{end_date.isoformat()}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/attendance/time-in/<int:employee_id>", methods=["POST"])
@admin_required
def admin_time_in(employee_id):
    employee = Employee.query.get_or_404(employee_id)
    attendance_date = parse_date(
        request.form.get("attendance_date")
    ) or date.today()

    record = Attendance.query.filter_by(
        employee_id=employee.id,
        attendance_date=attendance_date
    ).first()

    if not record:
        record = Attendance(
            employee_id=employee.id,
            attendance_date=attendance_date
        )
        db.session.add(record)

    if record.time_in:
        flash("Time In already exists. Use Edit to modify it.", "warning")
    else:
        record.time_in = datetime.now()
        calculate_attendance(record)
        db.session.commit()
        flash("Employee timed in.", "success")

    return redirect(url_for("attendance", date=attendance_date.isoformat()))


@app.route("/attendance/time-out/<int:employee_id>", methods=["POST"])
@admin_required
def admin_time_out(employee_id):
    employee = Employee.query.get_or_404(employee_id)
    attendance_date = parse_date(
        request.form.get("attendance_date")
    ) or date.today()

    record = Attendance.query.filter_by(
        employee_id=employee.id,
        attendance_date=attendance_date
    ).first()

    if not record:
        flash("No Time In record exists for this employee.", "danger")
        return redirect(
            url_for("attendance", date=attendance_date.isoformat())
        )

    record.time_out = datetime.now()
    calculate_attendance(record)
    db.session.commit()

    flash("Employee timed out.", "success")
    return redirect(url_for("attendance", date=attendance_date.isoformat()))


@app.route("/attendance/delete/<int:id>", methods=["POST"])
@admin_required
def delete_attendance(id):
    """Delete an attendance record manually by Admin.

    This removes only the selected attendance row. Existing generated payrolls
    are not silently recalculated; Admin should regenerate/review payroll when
    a deleted record affects an already-generated cutoff.
    """
    record = Attendance.query.get_or_404(id)
    employee = record.employee
    attendance_date = record.attendance_date
    employee_name = f"{employee.first_name} {employee.last_name}"
    leave_status = (record.status or "").upper()

    try:
        linked_leave = None
        if leave_status in ("VL", "SL"):
            # Attendance is the source of truth for leave credit usage.
            # If this attendance came from an approved leave request, repair
            # that request as well so the deleted date can be requested again.
            linked_leave = LeaveRequest.query.filter(
                LeaveRequest.employee_id == employee.id,
                LeaveRequest.leave_type == leave_status,
                LeaveRequest.status == "Approved",
                LeaveRequest.start_date <= attendance_date,
                LeaveRequest.end_date >= attendance_date,
            ).order_by(LeaveRequest.created_at.desc()).first()

        if linked_leave:
            # Preserve the approved request's remaining dates by cancelling
            # the original request and recreating the portions before/after
            # the deleted attendance date as separate approved requests.
            original_start = linked_leave.start_date
            original_end = linked_leave.end_date
            reason = linked_leave.reason
            remarks = linked_leave.admin_remarks or ""
            approved_by = linked_leave.approved_by
            approved_at = linked_leave.approved_at

            linked_leave.status = "Cancelled"
            linked_leave.admin_remarks = (
                (remarks + " | ") if remarks else ""
            ) + (
                f"Cancelled for {attendance_date.strftime('%b %d, %Y')} because the linked "
                "attendance record was deleted by Admin. Leave credit returned."
            )

            ranges = []
            before_end = attendance_date - timedelta(days=1)
            after_start = attendance_date + timedelta(days=1)
            if original_start <= before_end:
                ranges.append((original_start, before_end))
            if after_start <= original_end:
                ranges.append((after_start, original_end))

            for range_start, range_end in ranges:
                if leave_request_day_count(range_start, range_end) <= 0:
                    continue
                replacement = LeaveRequest(
                    employee_id=employee.id,
                    leave_type=leave_status,
                    start_date=range_start,
                    end_date=range_end,
                    reason=reason,
                    status="Approved",
                    admin_remarks=remarks or None,
                    approved_by=approved_by,
                    approved_at=approved_at,
                )
                db.session.add(replacement)

        db.session.delete(record)
        db.session.commit()

        if linked_leave:
            flash(
                f"{leave_status} attendance for {employee_name} on "
                f"{attendance_date.strftime('%b %d, %Y')} was deleted. "
                f"The {leave_status} credit for that date has been returned, "
                "and the leave request was adjusted.",
                "success"
            )
        elif leave_status in ("VL", "SL"):
            flash(
                f"{leave_status} attendance for {employee_name} on "
                f"{attendance_date.strftime('%b %d, %Y')} was deleted. "
                f"The {leave_status} credit has been returned.",
                "success"
            )
        else:
            flash(
                f"Attendance for {employee_name} on "
                f"{attendance_date.strftime('%b %d, %Y')} was deleted.",
                "success"
            )
    except Exception as exc:
        db.session.rollback()
        flash(f"Unable to delete attendance: {exc}", "danger")

    return redirect(url_for("attendance", date=attendance_date.isoformat()))


@app.route("/attendance/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_attendance(id):
    record = Attendance.query.get_or_404(id)

    if request.method == "POST":
        try:
            record.attendance_date = parse_date(
                request.form.get("attendance_date")
            ) or record.attendance_date
            requested_status = (request.form.get("status") or "Present").strip().upper()
            if requested_status in ("VL", "SL"):
                entitlement = float(getattr(record.employee, "vl_entitlement" if requested_status == "VL" else "sl_entitlement", 0) or 0)
                start = date(record.attendance_date.year, 1, 1)
                reg = getattr(record.employee, "regularization_date", None)
                if reg and reg > start:
                    start = reg
                used_before = Attendance.query.filter(
                    Attendance.employee_id == record.employee_id,
                    Attendance.attendance_date >= start,
                    Attendance.attendance_date <= date(record.attendance_date.year, 12, 31),
                    Attendance.status == requested_status,
                    Attendance.id != record.id
                ).count()
                if used_before + 1 > entitlement:
                    raise ValueError(f"{requested_status} entitlement exceeded. Entitled: {entitlement:g} day(s); already used: {used_before} day(s).")
                record.time_in = None
                record.time_out = None
                record.status = requested_status
                record.approved_overtime_hours = 0.0
                calculate_attendance(record)
                db.session.commit()
                flash(f"{requested_status} leave recorded.", "success")
                return redirect(url_for("attendance", date=record.attendance_date.isoformat()))

            record.time_in = parse_datetime_local(
                request.form.get("time_in")
            )
            record.time_out = parse_datetime_local(
                request.form.get("time_out")
            )

            if record.time_in and record.time_out and record.time_out < record.time_in:
                raise ValueError("Time Out cannot be earlier than Time In.")

            calculate_attendance(record)
            db.session.commit()

            flash("Attendance updated and recalculated.", "success")
            return redirect(
                url_for(
                    "attendance",
                    date=record.attendance_date.isoformat()
                )
            )
        except Exception as exc:
            db.session.rollback()
            flash(str(exc), "danger")

    return render_template("attendance_edit.html", record=record)


@app.route("/attendance/approve-ot/<int:id>", methods=["POST"])
@admin_required
def approve_attendance_ot(id):
    record = Attendance.query.get_or_404(id)

    try:
        requested = float(request.form.get("approved_overtime_hours") or 0)
    except ValueError:
        requested = 0.0

    requested = max(0.0, requested)
    available = max(0.0, float(record.overtime_hours or 0))

    # Admin can approve less than the available OT, but never more than actual OT.
    approved = min(requested, available)
    record.approved_overtime_hours = round(approved, 2)
    db.session.commit()

    if approved > 0:
        flash(f"Approved {approved:.2f} OT hour(s) for payroll.", "success")
    else:
        flash("OT approval cleared. OT will not be paid for this attendance.", "warning")

    return redirect(url_for("attendance", date=record.attendance_date.isoformat()))


@app.route("/employee/dashboard")
@employee_required
def employee_dashboard():
    user = User.query.get_or_404(session["user_id"])
    employee = user.employee
    today_attendance = Attendance.query.filter_by(
        employee_id=employee.id,
        attendance_date=date.today()
    ).first()

    latest_payroll = Payroll.query.filter_by(
        employee_id=employee.id
    ).order_by(Payroll.period_end.desc()).first()

    return render_template(
        "employee_dashboard.html",
        employee=employee,
        today_attendance=today_attendance,
        latest_payroll=latest_payroll
    )


@app.route("/employee/time-in", methods=["POST"])
@employee_required
def employee_time_in():
    user = User.query.get_or_404(session["user_id"])
    employee = user.employee
    today = date.today()

    record = Attendance.query.filter_by(
        employee_id=employee.id,
        attendance_date=today
    ).first()

    if not record:
        record = Attendance(
            employee_id=employee.id,
            attendance_date=today
        )
        db.session.add(record)

    if not record.time_in:
        record.time_in = datetime.now()
        calculate_attendance(record)
        db.session.commit()
        flash("Time In recorded.", "success")
    else:
        flash("You already timed in today.", "warning")

    return redirect(url_for("employee_dashboard"))


@app.route("/employee/time-out", methods=["POST"])
@employee_required
def employee_time_out():
    user = User.query.get_or_404(session["user_id"])
    employee = user.employee
    today = date.today()

    record = Attendance.query.filter_by(
        employee_id=employee.id,
        attendance_date=today
    ).first()

    if not record or not record.time_in:
        flash("You need to Time In first.", "danger")
    elif record.time_out:
        flash("You already timed out today.", "warning")
    else:
        record.time_out = datetime.now()
        calculate_attendance(record)
        db.session.commit()
        flash("Time Out recorded.", "success")

    return redirect(url_for("employee_dashboard"))


@app.route("/employee/attendance")
@employee_required
def employee_attendance():
    user = User.query.get_or_404(session["user_id"])
    records = Attendance.query.filter_by(
        employee_id=user.employee.id
    ).order_by(Attendance.attendance_date.desc()).all()

    return render_template(
        "employee_attendance.html",
        employee=user.employee,
        attendance=records
    )


@app.route("/employee/profile", methods=["GET", "POST"])
@employee_required
def employee_self_profile():
    user = User.query.get_or_404(session["user_id"])
    employee = user.employee

    if request.method == "POST":
        employee.email = request.form.get("email", "").strip() or None
        employee.phone = request.form.get("phone", "").strip() or None
        employee.address = request.form.get("address", "").strip() or None
        employee.emergency_name = request.form.get(
            "emergency_name", ""
        ).strip() or None
        employee.emergency_relationship = request.form.get(
            "emergency_relationship", ""
        ).strip() or None
        employee.emergency_phone = request.form.get(
            "emergency_phone", ""
        ).strip() or None
        employee.emergency_address = request.form.get(
            "emergency_address", ""
        ).strip() or None

        try:
            photo = request.files.get("profile_photo")
            if photo:
                save_profile_photo(employee, photo)

            db.session.commit()
            flash("Profile updated successfully.", "success")
        except Exception as exc:
            db.session.rollback()
            flash(str(exc), "danger")

        return redirect(url_for("employee_profile"))

    return render_template(
        "employee_self_profile.html",
        employee=employee
    )



LOAN_TYPES = (
    "SSS Salary Loan",
    "SSS Calamity Loan",
    "Pag-IBIG MPL",
    "Pag-IBIG Calamity Loan",
)

LOAN_FIELD_MAP = {
    "SSS Salary Loan": "loan_sss_salary",
    "SSS Calamity Loan": "loan_sss_calamity",
    "Pag-IBIG MPL": "loan_pagibig_mpl",
    "Pag-IBIG Calamity Loan": "loan_pagibig_calamity",
}


def get_active_loan_installments(employee_id, as_of_date=None):
    """Return active government loan records, grouped by loan type and total."""
    q = EmployeeLoan.query.filter(
        EmployeeLoan.employee_id == employee_id,
        EmployeeLoan.status == "Active",
        EmployeeLoan.loan_type.in_(LOAN_TYPES)
    ).order_by(EmployeeLoan.id.asc())

    if as_of_date:
        q = q.filter(or_(
            EmployeeLoan.start_date.is_(None),
            EmployeeLoan.start_date <= as_of_date
        ))
        q = q.filter(or_(
            EmployeeLoan.end_date.is_(None),
            EmployeeLoan.end_date >= as_of_date
        ))

    loans = q.all()
    breakdown = {field: 0.0 for field in LOAN_FIELD_MAP.values()}

    for loan in loans:
        balance = max(0.0, float(loan.balance or 0))
        installment = max(0.0, float(loan.deduction_per_payroll or 0))
        amount = min(balance, installment)
        field = LOAN_FIELD_MAP.get(loan.loan_type)
        if field:
            breakdown[field] += amount

    breakdown = {k: round(v, 2) for k, v in breakdown.items()}
    breakdown["total"] = round(sum(breakdown.values()), 2)
    return loans, breakdown


def allocate_loan_payment(payroll):
    """Allocate the payroll's government-loan deduction to the employee's loans by oldest first."""
    remaining = max(0.0, float(payroll.loan or 0))
    if remaining <= 0:
        return 0.0

    loans, _ = get_active_loan_installments(payroll.employee_id, payroll.period_end)
    allocated = 0.0

    for loan in loans:
        if remaining <= 0:
            break

        balance = max(0.0, float(loan.balance or 0))
        installment = max(0.0, float(loan.deduction_per_payroll or 0))
        amount = min(balance, installment, remaining)

        if amount <= 0:
            continue

        db.session.add(LoanPayment(
            loan_id=loan.id,
            payroll_id=payroll.id,
            amount=round(amount, 2)
        ))

        loan.balance = round(balance - amount, 2)
        if loan.balance <= 0.005:
            loan.balance = 0.0
            loan.status = "Paid"

        allocated += amount
        remaining -= amount

    return round(allocated, 2)



def is_leave_workday(d):
    """Leave consumes a credit only on scheduled Mon-Fri, non-holiday days."""
    if d.weekday() >= 5:  # Saturday/Sunday are not scheduled leave days.
        return False
    holiday = Holiday.query.filter_by(holiday_date=d, active=True).first()
    return holiday is None


def leave_request_day_count(start_date, end_date):
    if start_date > end_date:
        return 0
    total = 0
    current = start_date
    while current <= end_date:
        if is_leave_workday(current):
            total += 1
        current += timedelta(days=1)
    return total


def leave_balance_for_request(employee, leave_type, year, exclude_request_id=None):
    """Balance is entitlement less VL/SL already reflected in Attendance."""
    start = date(year, 1, 1)
    reg = getattr(employee, "regularization_date", None)
    if reg and reg > start:
        start = reg
    end = date(year, 12, 31)
    entitlement = float(getattr(employee, "vl_entitlement" if leave_type == "VL" else "sl_entitlement", 0) or 0)
    used = Attendance.query.filter(
        Attendance.employee_id == employee.id,
        Attendance.attendance_date >= start,
        Attendance.attendance_date <= end,
        Attendance.status == leave_type
    ).count()
    # Approved requests are mirrored into Attendance immediately, so Attendance
    # remains the single source of truth for used leave credits and we do not
    # double-count an approved request here.
    return max(0.0, entitlement - used)


app.jinja_env.globals["leave_request_day_count"] = leave_request_day_count

def request_overlaps_existing(employee_id, start_date, end_date, exclude_id=None):
    q = LeaveRequest.query.filter(
        LeaveRequest.employee_id == employee_id,
        LeaveRequest.status.in_(["Pending", "Approved"]),
        LeaveRequest.start_date <= end_date,
        LeaveRequest.end_date >= start_date,
    )
    if exclude_id:
        q = q.filter(LeaveRequest.id != exclude_id)
    return q.first()


@app.route("/leave-requests")
@admin_required
def admin_leave_requests():
    status = (request.args.get("status") or "Pending").strip()
    if status not in {"All", "Pending", "Approved", "Rejected"}:
        status = "Pending"
    query = LeaveRequest.query.join(Employee).order_by(LeaveRequest.created_at.desc())
    if status != "All":
        query = query.filter(LeaveRequest.status == status)
    requests_list = query.all()
    pending_count = LeaveRequest.query.filter_by(status="Pending").count()
    return render_template(
        "leave_requests_admin.html",
        leave_requests=requests_list,
        selected_status=status,
        pending_count=pending_count,
    )


@app.route("/leave-requests/<int:id>/approve", methods=["POST"])
@admin_required
def approve_leave_request(id):
    req = LeaveRequest.query.get_or_404(id)
    if req.status != "Pending":
        flash("Only Pending leave requests can be approved.", "warning")
        return redirect(url_for("admin_leave_requests"))

    employee = req.employee
    if req.start_date.year != req.end_date.year:
        flash("Leave requests cannot cross calendar years. Please submit separate requests.", "danger")
        return redirect(url_for("admin_leave_requests"))

    if employee.regularization_date and req.start_date < employee.regularization_date:
        flash("Leave cannot be approved before the employee's regularization date.", "danger")
        return redirect(url_for("admin_leave_requests"))

    requested_days = leave_request_day_count(req.start_date, req.end_date)
    if requested_days <= 0:
        flash("The selected dates contain no scheduled Mon-Fri leave day.", "danger")
        return redirect(url_for("admin_leave_requests"))

    overlap = request_overlaps_existing(employee.id, req.start_date, req.end_date, req.id)
    if overlap:
        flash(f"Overlapping {overlap.leave_type} request already exists ({overlap.status}).", "danger")
        return redirect(url_for("admin_leave_requests"))

    balance = leave_balance_for_request(employee, req.leave_type, req.start_date.year, req.id)
    if requested_days > balance:
        flash(
            f"Insufficient {req.leave_type} credit. Requested: {requested_days:g} day(s); available: {balance:g} day(s).",
            "danger"
        )
        return redirect(url_for("admin_leave_requests"))

    try:
        current = req.start_date
        while current <= req.end_date:
            if is_leave_workday(current):
                record = Attendance.query.filter_by(
                    employee_id=employee.id,
                    attendance_date=current
                ).first()
                if record and (record.time_in or record.time_out) and (record.status or "").upper() not in ("VL", "SL"):
                    raise ValueError(f"{current.strftime('%b %d, %Y')} already has attendance. Edit/clear that attendance before approving leave.")
                if not record:
                    record = Attendance(employee_id=employee.id, attendance_date=current)
                    db.session.add(record)
                record.time_in = None
                record.time_out = None
                record.total_hours = 0.0
                record.late_hours = 0.0
                record.undertime_hours = 0.0
                record.overtime_hours = 0.0
                record.approved_overtime_hours = 0.0
                record.status = req.leave_type
            current += timedelta(days=1)

        req.status = "Approved"
        req.admin_remarks = request.form.get("admin_remarks", "").strip() or None
        req.approved_by = session.get("user_id")
        req.approved_at = datetime.utcnow()
        db.session.commit()
        flash(
            f"{req.leave_type} request approved. Attendance was automatically updated and will be included as paid leave in payroll.",
            "success"
        )
    except Exception as exc:
        db.session.rollback()
        flash(str(exc), "danger")
    return redirect(url_for("admin_leave_requests"))


@app.route("/leave-requests/<int:id>/reject", methods=["POST"])
@admin_required
def reject_leave_request(id):
    req = LeaveRequest.query.get_or_404(id)
    if req.status != "Pending":
        flash("Only Pending leave requests can be rejected.", "warning")
        return redirect(url_for("admin_leave_requests"))
    req.status = "Rejected"
    req.admin_remarks = request.form.get("admin_remarks", "").strip() or None
    req.approved_by = session.get("user_id")
    req.approved_at = datetime.utcnow()
    db.session.commit()
    flash("Leave request rejected.", "success")
    return redirect(url_for("admin_leave_requests"))


@app.route("/employee/leave-requests", methods=["GET", "POST"])
@employee_required
def employee_leave_requests():
    user = User.query.get_or_404(session["user_id"])
    employee = user.employee
    today = date.today()

    if request.method == "POST":
        try:
            leave_type = (request.form.get("leave_type") or "VL").strip().upper()
            start_date = parse_date(request.form.get("start_date"))
            end_date = parse_date(request.form.get("end_date"))
            reason = request.form.get("reason", "").strip() or None

            if leave_type not in ("VL", "SL"):
                raise ValueError("Leave type must be VL or SL.")
            if not start_date or not end_date or start_date > end_date:
                raise ValueError("Please provide a valid leave date range.")
            if start_date.year != end_date.year:
                raise ValueError("Leave requests cannot cross calendar years. Submit separate requests.")
            if employee.regularization_date and start_date < employee.regularization_date:
                raise ValueError("You can only request leave starting on or after your regularization date.")

            requested_days = leave_request_day_count(start_date, end_date)
            if requested_days <= 0:
                raise ValueError("The selected dates contain no scheduled Mon-Fri leave day.")

            overlap = request_overlaps_existing(employee.id, start_date, end_date)
            if overlap:
                raise ValueError(f"You already have an overlapping {overlap.leave_type} request with status {overlap.status}.")

            balance = leave_balance_for_request(employee, leave_type, start_date.year)
            if requested_days > balance:
                raise ValueError(f"Insufficient {leave_type} credit. Requested: {requested_days:g} day(s); available: {balance:g} day(s).")

            db.session.add(LeaveRequest(
                employee_id=employee.id,
                leave_type=leave_type,
                start_date=start_date,
                end_date=end_date,
                reason=reason,
                status="Pending",
            ))
            db.session.commit()
            flash(f"{leave_type} request submitted for {requested_days:g} day(s). Waiting for Admin approval.", "success")
            return redirect(url_for("employee_leave_requests"))
        except Exception as exc:
            db.session.rollback()
            flash(str(exc), "danger")

    year = request.args.get("year", type=int) or today.year
    balances = {
        "vl": leave_balance_for_request(employee, "VL", year),
        "sl": leave_balance_for_request(employee, "SL", year),
        "vl_entitlement": float(employee.vl_entitlement or 0),
        "sl_entitlement": float(employee.sl_entitlement or 0),
    }
    requests_list = LeaveRequest.query.filter_by(employee_id=employee.id).order_by(LeaveRequest.created_at.desc()).all()
    return render_template(
        "leave_requests_employee.html",
        employee=employee,
        leave_requests=requests_list,
        balances=balances,
        year=year,
        today=today,
    )


@app.route("/loans")
@admin_required
def loans():
    loan_records = EmployeeLoan.query.join(Employee).order_by(
        EmployeeLoan.status.asc(),
        Employee.last_name.asc(),
        Employee.first_name.asc(),
        EmployeeLoan.id.asc()
    ).all()

    totals = {
        "original": round(sum(float(x.original_amount or 0) for x in loan_records), 2),
        "balance": round(sum(float(x.balance or 0) for x in loan_records), 2),
        "installment": round(sum(
            min(max(0.0, float(x.balance or 0)),
                max(0.0, float(x.deduction_per_payroll or 0)))
            for x in loan_records if x.status == "Active"
        ), 2),
    }
    return render_template("loans.html", loans=loan_records, totals=totals)


@app.route("/loans/add", methods=["GET", "POST"])
@admin_required
def add_loan():
    employees = Employee.query.order_by(
        Employee.last_name.asc(), Employee.first_name.asc()
    ).all()

    if request.method == "POST":
        original = max(0.0, float(request.form.get("original_amount") or 0))
        balance = max(0.0, float(request.form.get("balance") or original))
        balance = min(balance, original)
        installment = max(0.0, float(request.form.get("deduction_per_payroll") or 0))
        loan_type = request.form.get("loan_type", "SSS Salary Loan").strip()
        if loan_type not in LOAN_TYPES:
            flash("Invalid government loan type.", "danger")
            return render_template("loan_form.html", loan=locals().get("loan"), employees=employees)

        loan = EmployeeLoan(
            employee_id=int(request.form.get("employee_id")),
            loan_type=request.form.get("loan_type", "SSS Salary Loan").strip() or "SSS Salary Loan",
            reference_no=request.form.get("reference_no", "").strip() or None,
            original_amount=round(original, 2),
            balance=round(balance, 2),
            deduction_per_payroll=round(installment, 2),
            start_date=parse_date(request.form.get("start_date")),
            end_date=parse_date(request.form.get("end_date")),
            company_remittance=True,
            status="Active" if balance > 0 else "Paid",
            notes=request.form.get("notes", "").strip() or None
        )
        db.session.add(loan)
        db.session.commit()
        flash("Loan added successfully.", "success")
        return redirect(url_for("loans"))

    return render_template("loan_form.html", loan=None, employees=employees)


@app.route("/loans/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_loan(id):
    loan = EmployeeLoan.query.get_or_404(id)
    employees = Employee.query.order_by(
        Employee.last_name.asc(), Employee.first_name.asc()
    ).all()

    if request.method == "POST":
        original = max(0.0, float(request.form.get("original_amount") or 0))
        balance = max(0.0, float(request.form.get("balance") or 0))
        installment = max(0.0, float(request.form.get("deduction_per_payroll") or 0))
        loan_type = request.form.get("loan_type", "SSS Salary Loan").strip()
        if loan_type not in LOAN_TYPES:
            flash("Invalid government loan type.", "danger")
            return render_template("loan_form.html", loan=locals().get("loan"), employees=employees)

        loan.employee_id = int(request.form.get("employee_id"))
        loan.loan_type = request.form.get("loan_type", "Other").strip() or "Other"
        loan.reference_no = request.form.get("reference_no", "").strip() or None
        loan.original_amount = round(original, 2)
        loan.balance = round(min(balance, original), 2)
        loan.deduction_per_payroll = round(installment, 2)
        loan.start_date = parse_date(request.form.get("start_date"))
        loan.end_date = parse_date(request.form.get("end_date"))
        loan.company_remittance = request.form.get("company_remittance") == "yes"
        loan.status = "Active" if loan.balance > 0 else "Paid"
        loan.notes = request.form.get("notes", "").strip() or None

        db.session.commit()
        flash("Loan updated.", "success")
        return redirect(url_for("loans"))

    return render_template("loan_form.html", loan=loan, employees=employees)


@app.route("/loans/delete/<int:id>", methods=["POST"])
@admin_required
def delete_loan(id):
    loan = EmployeeLoan.query.get_or_404(id)
    if loan.payments:
        flash("Loan with payment history cannot be deleted.", "danger")
    else:
        db.session.delete(loan)
        db.session.commit()
        flash("Loan deleted.", "success")
    return redirect(url_for("loans"))


@app.route("/loans/report")
@admin_required
def loan_report():
    loans = EmployeeLoan.query.join(Employee).filter(
        EmployeeLoan.loan_type.in_(LOAN_TYPES)
    ).order_by(
        Employee.last_name.asc(), Employee.first_name.asc(),
        EmployeeLoan.loan_type.asc()
    ).all()

    rows = []
    for loan in loans:
        paid = round(sum(float(p.amount or 0) for p in loan.payments), 2)
        rows.append({
            "loan": loan,
            "paid": paid,
            "remaining": round(max(0.0, float(loan.balance or 0)), 2),
            "remittance_due": paid
        })

    type_totals = {}
    for loan in loans:
        type_totals.setdefault(loan.loan_type, {"deducted": 0.0, "balance": 0.0})
        type_totals[loan.loan_type]["balance"] += float(loan.balance or 0)
        type_totals[loan.loan_type]["deducted"] += sum(
            float(p.amount or 0) for p in loan.payments
        )

    for k in type_totals:
        type_totals[k] = {
            "deducted": round(type_totals[k]["deducted"], 2),
            "balance": round(type_totals[k]["balance"], 2)
        }

    totals = {
        "original": round(sum(float(r["loan"].original_amount or 0) for r in rows), 2),
        "balance": round(sum(r["remaining"] for r in rows), 2),
        "paid": round(sum(r["paid"] for r in rows), 2),
    }
    return render_template(
        "loan_report.html",
        rows=rows,
        totals=totals,
        type_totals=type_totals,
        loan_types=LOAN_TYPES
    )



# Unicode font for PDF payslips. Helvetica does not contain the Philippine peso sign (₱),
# which can render as a black square in generated PDF attachments.
PAYSLIP_FONT_PATH = os.path.join(BASE_DIR, "static", "fonts", "DejaVuSans.ttf")
PAYSLIP_FONT_BOLD_PATH = os.path.join(BASE_DIR, "static", "fonts", "DejaVuSans-Bold.ttf")
if os.path.exists(PAYSLIP_FONT_PATH):
    try:
        pdfmetrics.registerFont(TTFont("PayslipSans", PAYSLIP_FONT_PATH))
        if os.path.exists(PAYSLIP_FONT_BOLD_PATH):
            pdfmetrics.registerFont(TTFont("PayslipSans-Bold", PAYSLIP_FONT_BOLD_PATH))
    except Exception:
        pass

PDF_FONT = "PayslipSans" if "PayslipSans" in pdfmetrics.getRegisteredFontNames() else "Helvetica"
PDF_FONT_BOLD = "PayslipSans-Bold" if "PayslipSans-Bold" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold"

def _money(value):
    return f"₱{float(value or 0):,.2f}"


def build_payslip_pdf(payroll):
    """Build the emailed PDF in the same visual style as the on-screen payslip.

    The email attachment intentionally mirrors templates/payslip.html: A4 portrait,
    compact two-column payroll tables, company logo, signature block, and the
    diagonal SYSTEM GENERATED watermark.
    """
    buffer = BytesIO()

    def draw_watermark(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.Color(31/255, 68/255, 120/255, alpha=0.10))
        canvas.setFont(PDF_FONT_BOLD, 52)
        canvas.translate(A4[0] / 2, A4[1] / 2)
        canvas.rotate(38)
        canvas.drawCentredString(0, 0, "SYSTEM GENERATED")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=14*mm,
        leftMargin=14*mm,
        topMargin=12*mm,
        bottomMargin=12*mm,
        title=f"Payslip - {payroll.employee.employee_id}",
        author="UNITY PATH RECOVERY AND COLLECTION SERVICES OPC",
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="SlipTiny", parent=styles["Normal"], fontName=PDF_FONT,
        fontSize=7.4, leading=9.2, textColor=colors.HexColor("#111111")
    ))
    styles.add(ParagraphStyle(
        name="SlipSmall", parent=styles["Normal"], fontName=PDF_FONT,
        fontSize=8.2, leading=10.0, textColor=colors.HexColor("#111111")
    ))
    styles.add(ParagraphStyle(
        name="SlipSmallBold", parent=styles["Normal"], fontName=PDF_FONT_BOLD,
        fontSize=8.2, leading=10.0, textColor=colors.HexColor("#111111")
    ))
    styles.add(ParagraphStyle(
        name="SlipTitle", parent=styles["Normal"], fontName=PDF_FONT_BOLD,
        fontSize=16, leading=18, alignment=1, textColor=colors.HexColor("#111111")
    ))
    styles.add(ParagraphStyle(
        name="SlipPeriod", parent=styles["Normal"], fontName=PDF_FONT,
        fontSize=8.2, leading=10, alignment=1, textColor=colors.HexColor("#111111")
    ))
    styles.add(ParagraphStyle(
        name="SlipCenter", parent=styles["Normal"], fontName=PDF_FONT,
        fontSize=8, leading=10, alignment=1, textColor=colors.HexColor("#111111")
    ))

    story = []
    emp = payroll.employee
    # Use a Windows-compatible date format for the payslip header.
    period = f"{payroll.period_start.strftime('%B %d').replace(' 0', ' ')}-{payroll.period_end.strftime('%d').lstrip('0')}, {payroll.period_end.strftime('%Y')}"

    logo_path = os.path.join(BASE_DIR, "static", "logo.png")
    if os.path.exists(logo_path):
        from reportlab.platypus import Image
        logo = Image(logo_path, width=36*mm, height=12*mm, kind="proportional")
    else:
        logo = Paragraph("", styles["SlipSmall"])

    header = Table([
        [logo, Paragraph("<b>PAYSLIP</b><br/><font size='8'>Payslip For the period<br/><b>" + period + "</b></font>", styles["SlipTitle"])]
    ], colWidths=[45*mm, 135*mm])
    header.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LINEBELOW", (0,0), (-1,-1), 1.2, colors.HexColor("#1f4478")),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0),
    ]))
    story += [header, Spacer(1, 2.5*mm)]

    name = f"{emp.last_name}, {emp.first_name}{(' ' + emp.middle_name) if emp.middle_name else ''}"
    meta_data = [
        [Paragraph("Name", styles["SlipTiny"]), ":", Paragraph(f"<b>{name}</b>", styles["SlipSmall"])],
        [Paragraph("Position", styles["SlipTiny"]), ":", Paragraph(emp.position or "—", styles["SlipTiny"])],
        [Paragraph("Branch/Dept", styles["SlipTiny"]), ":", Paragraph(emp.department or "—", styles["SlipTiny"])],
        [Paragraph("Employee ID", styles["SlipTiny"]), ":", Paragraph(emp.employee_id or "—", styles["SlipTiny"])],
    ]
    meta = Table(meta_data, colWidths=[25*mm, 5*mm, 70*mm], hAlign="LEFT")
    meta.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0.5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0.5),
    ]))
    story += [meta, Spacer(1, 3*mm)]

    def money(value):
        return f"₱{float(value or 0):,.2f}"

    left_rows = [
        ["BASIC PAY", "Amount"],
        [f"Worked Days  <font size='6.5'>{payroll.worked_days or 0} day{'s' if (payroll.worked_days or 0) != 1 else ''}</font>", money((payroll.basic_pay or 0) - (payroll.vl_pay or 0) - (payroll.sl_pay or 0))],
        [f"VL with Pay  <font size='6.5'>{payroll.vl_days or 0} day{'s' if (payroll.vl_days or 0) != 1 else ''}</font>", money(payroll.vl_pay)],
        [f"SL with Pay  <font size='6.5'>{payroll.sl_days or 0} day{'s' if (payroll.sl_days or 0) != 1 else ''}</font>", money(payroll.sl_pay)],
        ["Adjustment Basic", money(payroll.adjustment_addition)],
        ["TOTAL BASIC", money((payroll.basic_pay or 0) + (payroll.adjustment_addition or 0))],
    ]
    allowance_total = (payroll.allowance or 0) + (payroll.incentive or 0) + (payroll.vl_conversion or 0) + (payroll.sl_conversion or 0)
    left_rows += [
        ["ALLOWANCES", "Amount"],
        ["Perfect Attendance", "—"],
        ["Transportation", money(payroll.allowance)],
        ["Meal Allowance", "—"],
        ["Birthday Bonus", "—"],
        ["Load Allowance", "—"],
        ["Lifestyle", "—"],
        ["Housing", "—"],
        ["Utility", "—"],
        ["Airfare", "—"],
        ["Other / Incentive", money(payroll.incentive)],
        ["VL Cash Conversion", money(payroll.vl_conversion)],
        ["SL Cash Conversion", money(payroll.sl_conversion)],
        ["TOTAL ALLOWANCE / OTHER EARNINGS", money(allowance_total)],
    ]

    premium_total = (payroll.holiday_pay or 0) + (payroll.rest_day_premium or 0) + (payroll.overtime_pay or 0)
    right_rows = [
        ["PREMIUMS & OVERTIME", "Amount"],
        ["Sunday Rest Day Premium", money(payroll.rest_day_premium)],
        ["Holiday Pay / Premium", money(payroll.holiday_pay)],
        [f"Regular OT  <font size='6.5'>{float(payroll.regular_ot_hours or 0):.2f} hrs</font>", money(payroll.regular_ot_pay)],
        [f"Holiday OT  <font size='6.5'>{float(payroll.holiday_ot_hours or 0):.2f} hrs</font>", money(payroll.holiday_ot_pay)],
        [f"Rest Day OT  <font size='6.5'>{float(payroll.rest_day_ot_hours or 0):.2f} hrs</font>", money(payroll.rest_day_ot_pay)],
        ["Adjustment OT", "—"],
        ["TOTAL PREMIUMS & OVERTIME", money(premium_total)],
        ["DEDUCTIONS", "Amount"],
        ["SSS Contribution", money(payroll.sss)],
        ["PhilHealth Contribution", money(payroll.philhealth)],
        ["HDMF Contribution", money(payroll.pagibig)],
        ["SSS Salary Loan", money(payroll.loan_sss_salary)],
        ["SSS Calamity Loan", money(payroll.loan_sss_calamity)],
        ["Pag-IBIG MPL", money(payroll.loan_pagibig_mpl)],
        ["Pag-IBIG Calamity Loan", money(payroll.loan_pagibig_calamity)],
        ["Tardy / Late Deduction", money(payroll.late_deduction)],
        ["Undertime Deduction", money(payroll.undertime_deduction)],
        ["Withholding Tax", money(payroll.withholding_tax)],
        ["Adjustment Deduction", money(payroll.adjustment_deduction)],
        ["Other Deduction", money(payroll.other_deduction)],
        ["TOTAL DEDUCTIONS", money(payroll.total_deductions)],
    ]

    def make_slip_table(rows):
        data = []
        for i, row in enumerate(rows):
            if i == 0 or row[0].startswith("TOTAL") or row[0] in ("ALLOWANCES", "DEDUCTIONS") or row[0] == "PREMIUMS & OVERTIME":
                a, b = Paragraph(f"<b>{row[0]}</b>", styles["SlipSmall"]), Paragraph(f"<b>{row[1]}</b>", styles["SlipSmall"])
            else:
                a, b = Paragraph(str(row[0]), styles["SlipTiny"]), Paragraph(str(row[1]), styles["SlipTiny"])
            data.append([a, b])
        t = Table(data, colWidths=[58*mm, 29*mm])
        ts = [
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (1,0), (1,-1), "RIGHT"),
            ("LEFTPADDING", (0,0), (-1,-1), 1),
            ("RIGHTPADDING", (0,0), (-1,-1), 1),
            ("TOPPADDING", (0,0), (-1,-1), 1.2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 1.2),
            ("LINEBELOW", (0,0), (-1,0), 1.0, colors.HexColor("#111111")),
        ]
        for i, row in enumerate(rows):
            if row[0] in ("ALLOWANCES", "DEDUCTIONS") or row[0] == "PREMIUMS & OVERTIME":
                ts += [("LINEBELOW", (0,i), (-1,i), 1.0, colors.HexColor("#111111"))]
            if str(row[0]).startswith("TOTAL"):
                ts += [("LINEABOVE", (0,i), (-1,i), 1.0, colors.HexColor("#111111"))]
                ts += [("FONTNAME", (0,i), (-1,i), "Helvetica-Bold")]
        t.setStyle(TableStyle(ts))
        return t

    columns = Table([[make_slip_table(left_rows), make_slip_table(right_rows)]], colWidths=[89*mm, 89*mm])
    columns.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0),
    ]))
    story += [columns, Spacer(1, 3*mm)]

    summary = Table([
        [Paragraph("GROSS PAY", styles["SlipSmall"]), Paragraph(f"<b>{money(payroll.gross_pay)}</b>", styles["SlipSmallBold"]),
         Paragraph("NET PAY", styles["SlipSmall"]), Paragraph(f"<b>{money(payroll.net_pay)}</b>", styles["SlipSmallBold"])]
    ], colWidths=[29*mm, 61*mm, 29*mm, 61*mm])
    summary.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (1,0), (1,0), "RIGHT"),
        ("ALIGN", (3,0), (3,0), "RIGHT"),
        ("LINEABOVE", (0,0), (1,0), 1.2, colors.HexColor("#111111")),
        ("LINEBELOW", (0,0), (1,0), 1.2, colors.HexColor("#111111")),
        ("BOX", (2,0), (3,0), 1.2, colors.HexColor("#111111")),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
    ]))
    story += [summary, Spacer(1, 6*mm)]

    status_line = payroll.status or "Draft"
    if payroll.pay_date:
        status_line += f" · Paid {payroll.pay_date.strftime('%B %d, %Y')}"
    ack = [
        Paragraph("I hereby acknowledge the receipt of the above amount:", styles["SlipCenter"]),
        Spacer(1, 9*mm),
        Paragraph(f"<b>{name}</b>", styles["SlipCenter"]),
        Paragraph("________________________________________", styles["SlipCenter"]),
        Paragraph(status_line, styles["SlipCenter"]),
    ]
    story += ack

    doc.build(story, onFirstPage=draw_watermark, onLaterPages=draw_watermark)
    buffer.seek(0)
    return buffer.getvalue()


def get_employee_email(employee):
    return (employee.company_email or employee.email or "").strip()


def send_payslip_email(payroll):
    recipient = get_employee_email(payroll.employee)
    if not recipient:
        raise ValueError("Employee has no Company Email or Personal Email.")

    host = os.getenv("MAIL_SERVER", "smtp.gmail.com")
    port = int(os.getenv("MAIL_PORT", "587"))
    username = os.getenv("MAIL_USERNAME", "").strip()
    password = os.getenv("MAIL_PASSWORD", "")
    sender = os.getenv("MAIL_FROM", username).strip()
    sender_name = os.getenv("MAIL_FROM_NAME", "UNITY PATH RECOVERY AND COLLECTION SERVICES OPC").strip()
    if not username or not password or not sender:
        raise ValueError("Email is not configured. Set MAIL_USERNAME, MAIL_PASSWORD and MAIL_FROM in .env.")

    period = f"{payroll.period_start.strftime('%B %d, %Y')} - {payroll.period_end.strftime('%B %d, %Y')}"
    employee_name = f"{payroll.employee.first_name} {payroll.employee.last_name}".strip()
    msg = EmailMessage()
    msg["Subject"] = f"Payslip – {period}"
    msg["From"] = f"{sender_name} <{sender}>"
    msg["To"] = recipient
    msg.set_content(
        f"Dear {employee_name},\n\n"
        f"Please find attached your payslip for the payroll period {period}.\n\n"
        f"Regards,\n{sender_name}\nHuman Resources"
    )
    pdf = build_payslip_pdf(payroll)
    safe_id = secure_filename(payroll.employee.employee_id or str(payroll.employee.id))
    filename = f"Payslip_{safe_id}_{payroll.period_start.isoformat()}_to_{payroll.period_end.isoformat()}.pdf"
    msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=filename)

    if port == 465:
        with smtplib.SMTP_SSL(host, port, timeout=30) as smtp:
            smtp.login(username, password)
            smtp.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=30) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(username, password)
            smtp.send_message(msg)


def migrate_v59_email_columns():
    """Add email delivery tracking fields to existing SQLite payrolls."""
    from sqlalchemy import text
    with db.engine.begin() as conn:
        existing = _table_columns("payrolls")
        columns = {
            "email_status": "VARCHAR(30) DEFAULT 'Not Yet Sent'",
            "email_sent_at": "DATETIME",
            "email_error": "TEXT",
        }
        for column, definition in columns.items():
            if column not in existing:
                conn.execute(text(f"ALTER TABLE payrolls ADD COLUMN {column} {definition}"))


@app.route("/payroll")
@admin_required
def payroll():
    payroll_records = Payroll.query.order_by(
        Payroll.period_end.desc(),
        Payroll.id.desc()
    ).all()
    return render_template("payroll.html", payroll_records=payroll_records)


@app.route("/payroll/<int:id>/send-email", methods=["POST"])
@admin_required
def send_payslip_email_route(id):
    payroll = Payroll.query.get_or_404(id)
    try:
        send_payslip_email(payroll)
        payroll.email_status = "Sent"
        payroll.email_sent_at = datetime.utcnow()
        payroll.email_error = None
        db.session.commit()
        flash(f"Payslip sent to {get_employee_email(payroll.employee)}.", "success")
    except Exception as exc:
        payroll.email_status = "Not Yet Sent"
        payroll.email_error = str(exc)[:1000]
        db.session.commit()
        flash(f"Payslip email failed: {exc}", "danger")
    return redirect(url_for("payroll"))


@app.route("/payroll/send-email-selected", methods=["POST"])
@admin_required
def send_selected_payslips_email():
    ids = request.form.getlist("payroll_ids")
    sent = 0
    failed = 0
    for raw_id in ids:
        try:
            payroll = Payroll.query.get(int(raw_id))
        except (TypeError, ValueError):
            payroll = None
        if payroll is None or payroll.email_status == "Sent":
            continue
        try:
            send_payslip_email(payroll)
            payroll.email_status = "Sent"
            payroll.email_sent_at = datetime.utcnow()
            payroll.email_error = None
            sent += 1
        except Exception as exc:
            payroll.email_status = "Not Yet Sent"
            payroll.email_error = str(exc)[:1000]
            failed += 1
    db.session.commit()
    if sent or failed:
        flash(f"Payslip email result: {sent} sent, {failed} failed.", "success" if not failed else "warning")
    else:
        flash("No Not Yet Sent payslips were selected.", "warning")
    return redirect(url_for("payroll"))


    payroll_records = Payroll.query.order_by(
        Payroll.period_end.desc(),
        Payroll.id.desc()
    ).all()

    return render_template(
        "payroll.html",
        payroll_records=payroll_records
    )


@app.route("/payroll/generate", methods=["GET", "POST"])
@admin_required
def generate_payroll():
    """Generate payroll by selecting a calendar month/year and cutoff.

    1st payroll  = day 1 through day 15, no employee SSS/PhilHealth/Pag-IBIG.
    2nd payroll = day 16 through the month's last day, with that month's
                  mandatory employee contributions and attendance calculations.

    This page intentionally has no future-cutoff blocking so Admin can test
    payroll generation with any month/cutoff.
    """
    settings = PayrollSettings.query.first()
    employees = Employee.query.filter_by(
        employment_status="Active"
    ).order_by(Employee.last_name.asc()).all()

    selected_year = request.form.get("year", type=int) if request.method == "POST" else None
    selected_month = request.form.get("month", type=int) if request.method == "POST" else None
    selected_payroll = request.form.get("payroll_type", "") if request.method == "POST" else ""
    pay_date = parse_date(request.form.get("pay_date")) if request.method == "POST" else None

    if request.method == "POST":
        if not selected_year or not selected_month or selected_month < 1 or selected_month > 12:
            flash("Please select a valid month and year.", "danger")
            return render_template(
                "payroll_generate.html",
                employees=employees,
                settings=settings,
                selected_year=selected_year,
                selected_month=selected_month,
                selected_payroll=selected_payroll,
                pay_date=pay_date,
            )

        if selected_payroll not in ("first", "second"):
            flash("Please choose 1st Payroll (1–15) or 2nd Payroll (16–last day).", "danger")
            return render_template(
                "payroll_generate.html",
                employees=employees,
                settings=settings,
                selected_year=selected_year,
                selected_month=selected_month,
                selected_payroll=selected_payroll,
                pay_date=pay_date,
            )

        first, mid, second, last = month_periods(selected_year, selected_month)
        if selected_payroll == "first":
            period_start, period_end = first, mid
        else:
            period_start, period_end = second, last

        # Pay Date is a payment-date record only. It does not determine the
        # attendance period or contribution month. For testing/convenience,
        # default it to the cutoff end if Admin leaves it blank.
        if pay_date is None:
            pay_date = period_end
        if pay_date < period_end:
            flash("Pay Date cannot be earlier than the end of the selected payroll period.", "danger")
            return render_template(
                "payroll_generate.html",
                employees=employees,
                settings=settings,
                selected_year=selected_year,
                selected_month=selected_month,
                selected_payroll=selected_payroll,
                pay_date=pay_date,
            )

        generated = 0
        skipped = 0

        for employee in employees:
            existing = Payroll.query.filter_by(
                employee_id=employee.id,
                period_start=period_start,
                period_end=period_end
            ).first()

            if existing:
                skipped += 1
                continue

            values = calculate_payroll_values(
                employee,
                period_start,
                period_end,
                settings
            )

            record = Payroll(
                employee_id=employee.id,
                period_start=period_start,
                period_end=period_end,
                pay_date=pay_date,
                **values
            )
            db.session.add(record)
            db.session.flush()

            # Only the 2nd payroll finalizes the month's mandatory
            # contribution obligation. The 1st payroll remains deduction-free.
            if selected_payroll == "second":
                upsert_monthly_contribution(
                    employee,
                    period_end,
                    settings,
                    payroll_id=record.id,
                    collected={
                        "sss": record.sss,
                        "philhealth": record.philhealth,
                        "pagibig": record.pagibig,
                    }
                )
            generated += 1

        db.session.commit()

        cutoff_label = "1st Payroll (1–15)" if selected_payroll == "first" else "2nd Payroll (16–last day)"
        contribution_note = (
            "No SSS/PhilHealth/Pag-IBIG employee deduction in this payroll."
            if selected_payroll == "first"
            else "Monthly SSS/PhilHealth/Pag-IBIG obligation for this month is finalized in this payroll."
        )
        flash(
            f"{cutoff_label} for {period_start.strftime('%B %Y')} generated for {generated} employees. "
            f"Skipped {skipped} existing records. {contribution_note}",
            "success"
        )
        return redirect(url_for("payroll"))

    return render_template(
        "payroll_generate.html",
        employees=employees,
        settings=settings,
        selected_year=selected_year,
        selected_month=selected_month,
        selected_payroll=selected_payroll,
        pay_date=pay_date,
    )


@app.route("/payroll/<int:id>/edit", methods=["GET", "POST"])
@admin_required
def edit_payroll(id):
    payroll = Payroll.query.get_or_404(id)

    if payroll.status == "Paid":
        flash("Paid payroll cannot be modified.", "danger")
        return redirect(url_for("view_payroll", id=id))

    employee = payroll.employee
    settings = PayrollSettings.query.first()

    if request.method == "POST":
        adjustment_addition = round(float(request.form.get("adjustment_addition") or 0), 2)
        adjustment_deduction = round(max(0.0, float(request.form.get("adjustment_deduction") or 0)), 2)
        adjustment_note = request.form.get("adjustment_note", "").strip() or None

        # Government loans are automatic from Loan Management and are not manually edited here.
        _, loan_breakdown = get_active_loan_installments(
            employee.id, payroll.period_end
        )
        loan = loan_breakdown["total"]
        other_deduction = round(max(0.0, float(request.form.get("other_deduction") or 0)), 2)

        # Recalculate earnings from attendance, while preserving daily rates.
        values = calculate_payroll_values(
            employee,
            payroll.period_start,
            payroll.period_end,
            settings,
            adjustment_addition,
            adjustment_deduction,
            adjustment_note
        )

        # Replace employee-specific deductions with the edited values and recalc total.
        values["loan"] = loan
        values["other_deduction"] = other_deduction

        # Government deductions must remain automatic. Re-run the statutory part
        # with the edited loan/other deduction values.
        deductions = calculate_deductions(
            employee,
            values["gross_pay"],
            settings,
            payroll.period_start,
            payroll.period_end,
            loan,
            other_deduction
        )
        values.update(deductions)

        if is_second_cutoff(payroll.period_start, payroll.period_end):
            upsert_monthly_contribution(
                employee, payroll.period_end, settings, payroll_id=payroll.id,
                collected={
                    "sss": values["sss"],
                    "philhealth": values["philhealth"],
                    "pagibig": values["pagibig"],
                }
            )

        values["loan_sss_salary"] = loan_breakdown["loan_sss_salary"]
        values["loan_sss_calamity"] = loan_breakdown["loan_sss_calamity"]
        values["loan_pagibig_mpl"] = loan_breakdown["loan_pagibig_mpl"]
        values["loan_pagibig_calamity"] = loan_breakdown["loan_pagibig_calamity"]

        # Explicitly persist attendance summary values on the payroll record.
        for field in (
            "worked_days", "regular_hours", "late_hours",
            "late_deduction", "undertime_hours", "undertime_deduction",
            "overtime_hours", "overtime_pay"
        ):
            setattr(payroll, field, values.get(field, 0))

        # Re-apply the one-time adjustment deduction as an additional deduction.
        # The adjustment_deduction is intentionally separate from "Other Deduction".
        total = (
            float(values["sss"] or 0) +
            float(values["philhealth"] or 0) +
            float(values["pagibig"] or 0) +
            float(values["withholding_tax"] or 0) +
            float(values.get("late_deduction") or 0) +
            float(values.get("undertime_deduction") or 0) +
            loan +
            other_deduction +
            adjustment_deduction
        )
        total = min(total, max(0.0, float(values["gross_pay"] or 0)))

        values["adjustment_addition"] = adjustment_addition
        values["adjustment_deduction"] = adjustment_deduction
        values["adjustment_note"] = adjustment_note
        values["total_deductions"] = round(total, 2)
        values["net_pay"] = round(
            float(values["gross_pay"] or 0) - total, 2
        )

        # Cash advance is retired from the UI/logic.
        values["cash_advance"] = 0.0

        for key, value in values.items():
            if hasattr(payroll, key):
                setattr(payroll, key, value)

        db.session.commit()
        flash("Individual payroll updated and recalculated.", "success")
        return redirect(url_for("view_payroll", id=id))

    return render_template("payroll_edit.html", payroll=payroll)


@app.route("/payroll/<int:id>")
@admin_required
def view_payroll(id):
    payroll = Payroll.query.get_or_404(id)
    contribution = get_monthly_contribution_record(
        payroll.employee_id, payroll.period_end.year, payroll.period_end.month
    ) if is_second_cutoff(payroll.period_start, payroll.period_end) else None
    return render_template("payroll_view.html", payroll=payroll, contribution=contribution)


@app.route("/payroll/<int:id>/approve", methods=["POST"])
@admin_required
def approve_payroll(id):
    payroll = Payroll.query.get_or_404(id)

    if payroll.status != "Draft":
        flash("Only Draft payroll can be approved.", "warning")
    else:
        payroll.status = "Approved"
        db.session.commit()
        flash("Payroll approved.", "success")

    return redirect(url_for("view_payroll", id=id))


@app.route("/payroll/<int:id>/paid", methods=["POST"])
@admin_required
def mark_payroll_paid(id):
    payroll = Payroll.query.get_or_404(id)

    if payroll.status != "Approved":
        flash("Only Approved payroll can be marked as Paid.", "warning")
    else:
        if LoanPayment.query.filter_by(payroll_id=payroll.id).count() == 0:
            allocate_loan_payment(payroll)

        if is_second_cutoff(payroll.period_start, payroll.period_end):
            settings = PayrollSettings.query.first()
            upsert_monthly_contribution(
                payroll.employee,
                payroll.period_end,
                settings,
                payroll_id=payroll.id,
                collected={
                    "sss": payroll.sss,
                    "philhealth": payroll.philhealth,
                    "pagibig": payroll.pagibig,
                }
            )

        payroll.status = "Paid"
        db.session.commit()
        flash("Payroll marked as Paid. Loan balance and monthly contribution collection updated.", "success")

    return redirect(url_for("view_payroll", id=id))



@app.route("/payroll/paid-selected", methods=["POST"])
@admin_required
def mark_selected_payrolls_paid():
    """Mark selected Approved payroll records as Paid in one action.

    Draft payrolls are not changed. For each Approved payroll, loan payments
    are allocated once and the linked monthly contribution is finalized for
    second-cutoff payrolls, matching the single-payroll Paid workflow.
    """
    selected_ids = request.form.getlist("payroll_ids")
    if not selected_ids:
        flash("Please select at least one Approved payroll to mark as Paid.", "warning")
        return redirect(url_for("payroll"))

    paid_count = 0
    skipped = 0
    invalid_ids = 0

    for raw_id in selected_ids:
        try:
            payroll_id = int(raw_id)
        except (TypeError, ValueError):
            invalid_ids += 1
            continue

        payroll = db.session.get(Payroll, payroll_id)
        if payroll is None:
            invalid_ids += 1
            continue

        if payroll.status != "Approved":
            skipped += 1
            continue

        if LoanPayment.query.filter_by(payroll_id=payroll.id).count() == 0:
            allocate_loan_payment(payroll)

        if is_second_cutoff(payroll.period_start, payroll.period_end):
            settings = PayrollSettings.query.first()
            upsert_monthly_contribution(
                payroll.employee,
                payroll.period_end,
                settings,
                payroll_id=payroll.id,
                collected={
                    "sss": payroll.sss,
                    "philhealth": payroll.philhealth,
                    "pagibig": payroll.pagibig,
                }
            )

        payroll.status = "Paid"
        paid_count += 1

    db.session.commit()

    parts = []
    if paid_count:
        parts.append(f"{paid_count} payroll record(s) marked as Paid.")
    if skipped:
        parts.append(f"{skipped} selected record(s) were skipped because they were not Approved.")
    if invalid_ids:
        parts.append(f"{invalid_ids} invalid/missing selection(s) were ignored.")

    flash(" ".join(parts), "success" if paid_count else "warning")
    return redirect(url_for("payroll"))


def remove_financial_links_for_payroll(payroll):
    """Remove financial rows whose source is the payroll being deleted.

    Salary is a one-to-one expense. Government contribution expenses are
    monthly aggregates, so they are rebuilt after deletion and removed only
    when no valid second-cutoff payroll remains for that month.
    """
    salary = Expense.query.filter_by(
        source_type="payroll_salary",
        source_key=f"payroll_salary:{payroll.id}"
    ).first()
    if salary is not None:
        db.session.delete(salary)

    month_key = None
    if payroll.period_end:
        month_key = f"{payroll.period_end.year}-{payroll.period_end.month:02d}"
    contribution = MonthlyContribution.query.filter_by(source_payroll_id=payroll.id).first()
    if contribution is not None:
        db.session.delete(contribution)

    return month_key


@app.route("/payroll/delete-selected", methods=["POST"])
@admin_required
def delete_selected_payroll():
    """Delete selected payroll records, including Paid records.

    When a Paid payroll is deleted, reverse its linked loan payments so the
    employee loan balances return to the state before that payroll was paid.
    Remove the linked monthly contribution record as well to avoid stale
    contribution data. Attendance and employee records are never deleted.
    """
    selected_ids = request.form.getlist("payroll_ids")
    if not selected_ids:
        flash("Please select at least one payroll to delete.", "warning")
        return redirect(url_for("payroll"))

    deleted = 0
    invalid_ids = 0

    for raw_id in selected_ids:
        try:
            payroll_id = int(raw_id)
        except (TypeError, ValueError):
            invalid_ids += 1
            continue

        payroll = db.session.get(Payroll, payroll_id)
        if payroll is None:
            invalid_ids += 1
            continue

        # Reverse any loan balances reduced by this payroll before removing
        # the payment records.
        for loan_payment in LoanPayment.query.filter_by(payroll_id=payroll.id).all():
            loan = db.session.get(EmployeeLoan, loan_payment.loan_id)
            if loan is not None:
                loan.balance = round(
                    min(float(loan.original_amount or 0),
                        max(0.0, float(loan.balance or 0) + float(loan_payment.amount or 0))),
                    2
                )
                loan.status = "Active" if loan.balance > 0.005 else "Paid"
            db.session.delete(loan_payment)

        # Remove all financial rows sourced by this payroll.
        remove_financial_links_for_payroll(payroll)
        db.session.delete(payroll)
        deleted += 1

    db.session.commit()

    # Global cleanup removes any orphaned legacy rows left by older versions.
    cleanup_orphan_financial_records()

    parts = []
    if deleted:
        parts.append(f"{deleted} payroll record(s) deleted successfully.")
    if invalid_ids:
        parts.append(f"{invalid_ids} invalid/missing selection(s) were ignored.")

    flash(" ".join(parts), "success" if deleted else "warning")
    return redirect(url_for("payroll"))


@app.route("/payroll/<int:id>/delete", methods=["POST"])
@admin_required
def delete_payroll(id):
    """Delete one payroll record, including Paid records."""
    payroll = Payroll.query.get_or_404(id)

    # Reverse loan payments first so deleting a Paid payroll does not leave
    # the employee's loan balance permanently reduced.
    for loan_payment in LoanPayment.query.filter_by(payroll_id=payroll.id).all():
        loan = db.session.get(EmployeeLoan, loan_payment.loan_id)
        if loan is not None:
            loan.balance = round(
                min(float(loan.original_amount or 0),
                    max(0.0, float(loan.balance or 0) + float(loan_payment.amount or 0))),
                2
            )
            loan.status = "Active" if loan.balance > 0.005 else "Paid"
        db.session.delete(loan_payment)

    # Remove all financial rows sourced by this payroll.
    remove_financial_links_for_payroll(payroll)
    db.session.delete(payroll)
    db.session.commit()
    cleanup_orphan_financial_records()
    flash("Payroll deleted successfully.", "success")
    return redirect(url_for("payroll"))


@app.route("/payroll/monthly-contributions")
@admin_required
def monthly_contributions():
    today = date.today()
    try:
        year = int(request.args.get("year", today.year))
        month = int(request.args.get("month", today.month))
    except (TypeError, ValueError):
        year, month = today.year, today.month

    rows = MonthlyContribution.query.filter_by(
        contribution_year=year,
        contribution_month=month
    ).join(Employee).order_by(Employee.last_name.asc(), Employee.first_name.asc()).all()

    totals = {key: 0.0 for key in (
        "sss_due", "philhealth_due", "pagibig_due",
        "sss_collected", "philhealth_collected", "pagibig_collected",
        "sss_uncollected", "philhealth_uncollected", "pagibig_uncollected",
        "sss_employer", "philhealth_employer", "pagibig_employer"
    )}
    for row in rows:
        for key in totals:
            totals[key] += float(getattr(row, key) or 0)
    totals = {k: round(v, 2) for k, v in totals.items()}
    totals["employee_due_total"] = round(sum(totals[k] for k in ("sss_due", "philhealth_due", "pagibig_due")), 2)
    totals["employee_collected_total"] = round(sum(totals[k] for k in ("sss_collected", "philhealth_collected", "pagibig_collected")), 2)
    totals["employee_uncollected_total"] = round(sum(totals[k] for k in ("sss_uncollected", "philhealth_uncollected", "pagibig_uncollected")), 2)
    totals["company_total"] = round(sum(totals[k] for k in ("sss_employer", "philhealth_employer", "pagibig_employer")), 2)

    return render_template(
        "monthly_contributions.html",
        rows=rows,
        totals=totals,
        year=year,
        month=month
    )


def purge_month_without_payroll(year, month):
    """Hard source-of-truth guard: no 2nd payroll means no monthly shares.

    This handles legacy/stale rows even when they have no source_payroll_id or
    point to a deleted/invalid payroll. It is safe to call before rendering
    Contribution Shares and the Financial Summary.
    """
    start = date(year, month, 16)
    end = date(year, month, calendar.monthrange(year, month)[1])
    valid_payroll_ids = {p.id for p in Payroll.query.filter(
        Payroll.status.in_(["Approved", "Paid"]),
        Payroll.period_start == start,
        Payroll.period_end == end,
    ).all()}

    records = MonthlyContribution.query.filter_by(
        contribution_year=year, contribution_month=month
    ).all()
    for record in records:
        if record.source_payroll_id not in valid_payroll_ids:
            db.session.delete(record)

    # Remove auto-generated government expense rows when there is no valid
    # contribution source. A manually paid row is also stale if its payroll
    # source no longer exists, so it must be removed rather than kept forever.
    if not valid_payroll_ids:
        for expense in Expense.query.filter_by(source_type="payroll_contribution").all():
            if expense.source_key and expense.source_key.startswith(f"{year}-{month:02d}:"):
                db.session.delete(expense)

    db.session.commit()
    return valid_payroll_ids


@app.route("/payroll/contributions")
@admin_required
def payroll_contributions():
    """Show statutory shares for a selected contribution month.

    When a second-cutoff payroll exists, the monthly contribution record is
    the source of truth, so Employee Share Due/Collected always reconciles
    with the payroll's actual SSS/PhilHealth/Pag-IBIG lines.
    """
    today = date.today()
    try:
        year = int(request.args.get("year", today.year))
        month = int(request.args.get("month", today.month))
    except (TypeError, ValueError):
        year, month = today.year, today.month
    if month < 1 or month > 12:
        month = today.month

    # Payroll is the source of truth for statutory contributions. Clean any
    # orphaned contribution records first, then show only employees who have
    # an actual finalized (2nd-cutoff) payroll contribution record for this
    # month. Do NOT calculate contributions for every active employee when no
    # payroll exists; that would create phantom obligations.
    # Ensure the selected month cannot display phantom shares after payroll
    # deletion. Payroll is the only source of monthly statutory obligations.
    purge_month_without_payroll(year, month)
    cleanup_orphan_financial_records(year, month)
    settings = PayrollSettings.query.first()

    totals = {
        "sss_employee": 0.0, "sss_employer": 0.0,
        "philhealth_employee": 0.0, "philhealth_employer": 0.0,
        "pagibig_employee": 0.0, "pagibig_employer": 0.0,
        "sss_collected": 0.0, "philhealth_collected": 0.0, "pagibig_collected": 0.0,
        "employee_due": 0.0, "employee_collected": 0.0, "employee_uncollected": 0.0,
    }
    rows = []

    # Only finalized monthly contribution records are displayed. These records
    # are created by the 2nd payroll. If the payroll is deleted, its linked
    # MonthlyContribution is deleted by the payroll delete route and by the
    # orphan cleanup above, so this page becomes empty automatically.
    records = (MonthlyContribution.query
               .join(Payroll, MonthlyContribution.source_payroll_id == Payroll.id)
               .join(Employee)
               .filter(
                   MonthlyContribution.contribution_year == year,
                   MonthlyContribution.contribution_month == month,
                   Payroll.status.in_(["Approved", "Paid"]),
                   Payroll.period_start == date(year, month, 16),
                   Payroll.period_end == date(year, month, calendar.monthrange(year, month)[1]),
               )
               .order_by(Employee.last_name.asc(), Employee.first_name.asc()).all())

    for record in records:
        employee = record.employee
        row = {
            "employee": employee,
            "sss_remuneration": float(record.sss_remuneration or 0),
            "sss_msc": float(record.sss_msc or 0),
            "philhealth_basic_salary": float(record.philhealth_basic_salary or 0),
            "pagibig_monthly_compensation": float(record.pagibig_monthly_compensation or 0),
            "sss_employee": float(record.sss_due or 0),
            "sss_employer": float(record.sss_employer or 0),
            "philhealth_employee": float(record.philhealth_due or 0),
            "philhealth_employer": float(record.philhealth_employer or 0),
            "pagibig_employee": float(record.pagibig_due or 0),
            "pagibig_employer": float(record.pagibig_employer or 0),
            "sss_collected": float(record.sss_collected or 0),
            "philhealth_collected": float(record.philhealth_collected or 0),
            "pagibig_collected": float(record.pagibig_collected or 0),
            "sss_uncollected": float(record.sss_uncollected or 0),
            "philhealth_uncollected": float(record.philhealth_uncollected or 0),
            "pagibig_uncollected": float(record.pagibig_uncollected or 0),
            "status": record.status,
        }

        row["employee_due"] = round(row["sss_employee"] + row["philhealth_employee"] + row["pagibig_employee"], 2)
        row["employee_collected"] = round(row["sss_collected"] + row["philhealth_collected"] + row["pagibig_collected"], 2)
        row["employee_uncollected"] = round(row["sss_uncollected"] + row["philhealth_uncollected"] + row["pagibig_uncollected"], 2)
        rows.append(row)

        for key in totals:
            totals[key] += float(row.get(key, 0) or 0)

    totals = {k: round(v, 2) for k, v in totals.items()}
    totals["employee_total"] = totals["employee_due"]
    totals["company_total"] = round(totals["sss_employer"] + totals["philhealth_employer"] + totals["pagibig_employer"], 2)
    totals["combined_total"] = round(totals["employee_due"] + totals["company_total"], 2)

    return render_template(
        "payroll_contributions.html",
        rows=rows, totals=totals, settings=settings, year=year, month=month
    )


@app.route("/payroll/final-pay/<int:id>", methods=["GET", "POST"])
@admin_required
def final_pay(id):
    employee = Employee.query.get_or_404(id)
    settings = PayrollSettings.query.first()

    if employee.employment_status == "Active":
        flash("Set the employee to a separated status before generating Final Pay.", "warning")
        return redirect(url_for("employee_profile", id=id))

    separation_date = employee.separation_date
    if request.method == "POST":
        separation_date = parse_date(request.form.get("separation_date"))
        separation_type = request.form.get("separation_type", "Resignation").strip()
        separation_reason = request.form.get("separation_reason", "").strip() or None

        if not separation_date:
            flash("Separation date is required.", "danger")
            return redirect(url_for("final_pay", id=id))

        employee.separation_date = separation_date
        employee.separation_type = separation_type
        employee.separation_reason = separation_reason
        employee.employment_status = "Resigned" if separation_type == "Resignation" else "Separated"

        values = calculate_final_pay_values(employee, separation_date, settings)

        existing = FinalPay.query.filter_by(employee_id=employee.id).first()
        if existing:
            for key, value in values.items():
                if hasattr(existing, key) and key not in {"last_payroll_end", "coverage_start", "coverage_end"}:
                    setattr(existing, key, value)
            existing.separation_date = separation_date
            existing.separation_type = separation_type
            existing.status = "Draft"
            record = existing
        else:
            record = FinalPay(
                employee_id=employee.id,
                separation_date=separation_date,
                separation_type=separation_type,
                status="Draft",
                **{k: v for k, v in values.items() if hasattr(FinalPay, k)}
            )
            db.session.add(record)

        db.session.commit()
        flash("Final Pay generated.", "success")
        return redirect(url_for("view_final_pay", id=record.id))

    return render_template(
        "final_pay_generate.html",
        employee=employee,
        settings=settings
    )


@app.route("/final-pay/<int:id>")
@admin_required
def view_final_pay(id):
    record = FinalPay.query.get_or_404(id)
    return render_template("final_pay_view.html", final_pay=record)


@app.route("/final-pay/<int:id>/approve", methods=["POST"])
@admin_required
def approve_final_pay(id):
    record = FinalPay.query.get_or_404(id)
    if record.status != "Draft":
        flash("Only Draft Final Pay can be approved.", "warning")
    else:
        record.status = "Approved"
        db.session.commit()
        flash("Final Pay approved.", "success")
    return redirect(url_for("view_final_pay", id=id))


@app.route("/final-pay/<int:id>/paid", methods=["POST"])
@admin_required
def paid_final_pay(id):
    record = FinalPay.query.get_or_404(id)
    if record.status != "Approved":
        flash("Only Approved Final Pay can be marked as Paid.", "warning")
    else:
        record.status = "Paid"
        db.session.commit()
        flash("Final Pay marked as Paid.", "success")
    return redirect(url_for("view_final_pay", id=id))



def sync_thirteenth_month_records(year):
    """Create/update one persistent 13th-month tracker row per employee/year.

    The amount is always recalculated from the same basic-salary source used by
    the 13th Month report. Payment status is intentionally preserved so that
    refreshing/recalculating payroll does not erase a prior Paid status.
    """
    employees = Employee.query.order_by(Employee.last_name.asc(), Employee.first_name.asc()).all()
    year_start = date(year, 1, 1)
    year_end = date(year, 12, 31)
    changed = False
    for employee in employees:
        separation_date = employee.separation_date
        separated_in_year = (
            separation_date is not None
            and year_start <= separation_date <= year_end
            and (employee.employment_status or "").strip().lower() in ("resigned", "separated", "inactive")
        )
        through_date = separation_date if separated_in_year else None
        basic = calculate_ytd_basic_pay(employee.id, year, through_date=through_date)

        # Match the existing report's Final Pay treatment for unpaid basic pay
        # through separation. Avoid adding it twice if it is already represented
        # by an approved/paid payroll.
        if separated_in_year:
            final = FinalPay.query.filter(
                FinalPay.employee_id == employee.id,
                FinalPay.separation_date == separation_date,
                FinalPay.separation_date >= year_start,
                FinalPay.separation_date <= year_end
            ).order_by(FinalPay.id.desc()).first()
            if final:
                basic += float(final.basic_pay or 0)
        basic = round(basic, 2)
        amount = round(basic / 12.0, 2)
        due_date = (separation_date + timedelta(days=30)) if separated_in_year else date(year, 12, 24)
        rec = ThirteenthMonthRecord.query.filter_by(employee_id=employee.id, year=year).first()
        if rec is None:
            rec = ThirteenthMonthRecord(employee_id=employee.id, year=year, status="Pending")
            db.session.add(rec)
            changed = True
        if (round(float(rec.basic_salary_earned or 0),2) != basic or
            round(float(rec.thirteenth_amount or 0),2) != amount or
            rec.due_date != due_date):
            rec.basic_salary_earned = basic
            rec.thirteenth_amount = amount
            rec.due_date = due_date
            changed = True
        rec.remarks = "Separated – Pro-rated" if separated_in_year else "Regular Year-end"
    if changed:
        db.session.commit()


def sync_thirteenth_month_expenses(year):
    """Mirror 13th-month tracker rows into company expense records."""
    records = ThirteenthMonthRecord.query.filter_by(year=year).all()
    for rec in records:
        source_key = f"{year}:{rec.id}"
        expense = Expense.query.filter_by(source_type="thirteenth_month", source_key=source_key).first()
        if expense is None:
            expense = Expense(
                expense_date=rec.paid_date or rec.due_date or date(year,12,24),
                name=f"13th Month - {rec.employee.first_name} {rec.employee.last_name}",
                category="13th Month Pay",
                amount=round(float(rec.thirteenth_amount or 0),2),
                due_date=rec.due_date,
                paid_date=rec.paid_date if rec.status == "Paid" else None,
                status=rec.status,
                remarks=f"Auto-generated from 13th Month Tracker {year}.",
                source_type="thirteenth_month",
                source_key=source_key,
                payment_reference=rec.payment_reference,
            )
            db.session.add(expense)
        else:
            expense.amount = round(float(rec.thirteenth_amount or 0),2)
            expense.name = f"13th Month - {rec.employee.first_name} {rec.employee.last_name}"
            expense.category = "13th Month Pay"
            expense.due_date = rec.due_date
            expense.status = rec.status
            expense.paid_date = rec.paid_date if rec.status == "Paid" else None
            expense.payment_reference = rec.payment_reference
            expense.remarks = f"Auto-generated from 13th Month Tracker {year}."
            if rec.status == "Paid" and rec.paid_date:
                expense.expense_date = rec.paid_date
            elif rec.due_date:
                expense.expense_date = rec.due_date
    db.session.commit()


@app.route("/payroll/13th-month", methods=["GET", "POST"])
@admin_required
def thirteenth_month_report():
    today = date.today()
    year = request.args.get("year", type=int) or request.form.get("year", type=int) or today.year
    if year < 2000 or year > 2100:
        year = today.year

    if request.method == "POST":
        action = request.form.get("action", "")
        ids = []
        for raw in request.form.getlist("record_ids"):
            try:
                ids.append(int(raw))
            except (TypeError, ValueError):
                pass
        selected = ThirteenthMonthRecord.query.filter(
            ThirteenthMonthRecord.id.in_(ids), ThirteenthMonthRecord.year == year
        ).all() if ids else []
        if action in ("mark_paid", "mark_pending"):
            paid_date = parse_date(request.form.get("paid_date")) or today
            ref = (request.form.get("payment_reference") or "").strip()[:150]
            for rec in selected:
                if action == "mark_paid":
                    rec.status = "Paid"
                    rec.paid_date = paid_date
                    if ref:
                        rec.payment_reference = ref
                else:
                    rec.status = "Pending"
                    rec.paid_date = None
            db.session.commit()
            sync_thirteenth_month_expenses(year)
            flash(f"{len(selected)} 13th month record(s) updated.", "success")
            return redirect(url_for("thirteenth_month_report", year=year))

    sync_thirteenth_month_records(year)
    sync_thirteenth_month_expenses(year)
    records = ThirteenthMonthRecord.query.filter_by(year=year).join(Employee).order_by(Employee.last_name.asc(), Employee.first_name.asc()).all()
    total_basic = round(sum(float(r.basic_salary_earned or 0) for r in records), 2)
    total_thirteenth = round(sum(float(r.thirteenth_amount or 0) for r in records), 2)
    paid_total = round(sum(float(r.thirteenth_amount or 0) for r in records if r.status == "Paid"), 2)
    pending_total = round(sum(float(r.thirteenth_amount or 0) for r in records if r.status == "Pending"), 2)
    return render_template("thirteenth_month.html", records=records, year=year,
                           total_basic=total_basic, total_thirteenth=total_thirteenth,
                           paid_total=paid_total, pending_total=pending_total, today=today)


@app.route("/payroll/13th-month/export")
@admin_required
def export_thirteenth_month():
    year = request.args.get("year", type=int) or date.today().year
    sync_thirteenth_month_records(year)
    records = ThirteenthMonthRecord.query.filter_by(year=year).join(Employee).order_by(Employee.last_name.asc(), Employee.first_name.asc()).all()
    wb = Workbook(); ws = wb.active; ws.title = "13th Month"
    ws.append([f"13th Month Pay Tracker - {year}"])
    ws.append([])
    ws.append(["Employee Name", "Employee ID", "Basic Salary Earned", "13th Month", "Status", "Due Date", "Paid Date", "Payment Reference"])
    for c in ws[3]: c.font = Font(bold=True)
    for r in records:
        ws.append([f"{r.employee.first_name} {r.employee.last_name}".strip(), r.employee.employee_id,
                   r.basic_salary_earned or 0, r.thirteenth_amount or 0, r.status,
                   r.due_date, r.paid_date, r.payment_reference or ""])
    for row in ws.iter_rows(min_row=4, min_col=3, max_col=4):
        for c in row: c.number_format = '₱#,##0.00'
    ws.freeze_panes = "A4"
    for i,w in enumerate([30,18,22,18,15,15,15,25],1): ws.column_dimensions[get_column_letter(i)].width=w
    out=BytesIO(); wb.save(out); out.seek(0)
    return send_file(out, as_attachment=True, download_name=f"13th_month_tracker_{year}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/payroll/settings", methods=["GET", "POST"])
@admin_required
def payroll_settings():
    settings = PayrollSettings.query.first()

    if request.method == "POST":
        settings.working_days_per_month = float(
            request.form.get("working_days_per_month") or 26
        )
        settings.hours_per_day = float(
            request.form.get("hours_per_day") or 8
        )
        settings.overtime_multiplier = float(
            request.form.get("overtime_multiplier") or 1.25
        )
        settings.sss_employee_amount = float(
            request.form.get("sss_employee_amount") or 0
        )
        settings.philhealth_rate = float(
            request.form.get("philhealth_rate") or 0
        )
        settings.pagibig_rate = float(
            request.form.get("pagibig_rate") or 0
        )
        settings.pagibig_max_employee_contribution = float(
            request.form.get("pagibig_max_employee_contribution") or 0
        )
        settings.withholding_tax_rate = float(
            request.form.get("withholding_tax_rate") or 0
        )

        db.session.commit()
        flash("Payroll settings saved.", "success")
        return redirect(url_for("payroll_settings"))

    return render_template(
        "payroll_settings.html",
        settings=settings
    )


@app.route("/employee/payroll")
@employee_required
def employee_payroll():
    user = User.query.get_or_404(session["user_id"])
    records = Payroll.query.filter_by(
        employee_id=user.employee.id
    ).order_by(Payroll.period_end.desc()).all()

    return render_template(
        "employee_payroll.html",
        payroll_records=records,
        employee=user.employee
    )


@app.route("/payroll/<int:id>/payslip")
@admin_required
def admin_payslip(id):
    payroll = Payroll.query.get_or_404(id)
    return render_template("payslip.html", payroll=payroll)


@app.route("/employee/payroll/<int:id>/payslip")
@employee_required
def employee_payslip(id):
    user = User.query.get_or_404(session["user_id"])
    payroll = Payroll.query.get_or_404(id)

    if payroll.employee_id != user.employee.id:
        abort(403)

    return render_template("payslip.html", payroll=payroll)


def _table_columns(table_name):
    """Return column names for a table using SQLAlchemy inspection (SQLite/PostgreSQL)."""
    inspector = inspect(db.engine)
    if table_name not in inspector.get_table_names():
        return set()
    return {col["name"] for col in inspector.get_columns(table_name)}


def migrate_v12_columns():
    """Add V12 columns to an existing V10/V11 SQLite database without losing data."""
    from sqlalchemy import text

    with db.engine.begin() as conn:
        tables = {
            "employees": [
                ("daily_allowance", "FLOAT DEFAULT 0"),
                ("daily_incentive", "FLOAT DEFAULT 0"),
                ("loan_deduction", "FLOAT DEFAULT 0"),
            ],
            "payrolls": [
                ("adjustment_addition", "FLOAT DEFAULT 0"),
                ("adjustment_deduction", "FLOAT DEFAULT 0"),
                ("adjustment_note", "TEXT"),
            ],
        }

        for table, columns in tables.items():
            existing = _table_columns(table)
            for column, definition in columns:
                if column not in existing:
                    conn.execute(
                        text(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
                    )


def migrate_v15_columns():
    """Add V15 payroll breakdown columns to an existing SQLite database."""
    from sqlalchemy import text
    with db.engine.begin() as conn:
        columns = [
            ("worked_days", "INTEGER DEFAULT 0"),
            ("regular_hours", "FLOAT DEFAULT 0"),
            ("late_hours", "FLOAT DEFAULT 0"),
            ("undertime_hours", "FLOAT DEFAULT 0"),
            ("undertime_deduction", "FLOAT DEFAULT 0"),
            ("loan_sss_salary", "FLOAT DEFAULT 0"),
            ("loan_sss_calamity", "FLOAT DEFAULT 0"),
            ("loan_pagibig_mpl", "FLOAT DEFAULT 0"),
            ("loan_pagibig_calamity", "FLOAT DEFAULT 0"),
        ]
        existing = _table_columns("payrolls")
        for col, definition in columns:
            if col not in existing:
                conn.execute(text(f"ALTER TABLE payrolls ADD COLUMN {col} {definition}"))

def migrate_v16_attendance_columns():
    """Add V16 approved OT field to existing attendance tables."""
    from sqlalchemy import text
    with db.engine.begin() as conn:
        existing = _table_columns("attendances")
        if "approved_overtime_hours" not in existing:
            conn.execute(text(
                "ALTER TABLE attendances ADD COLUMN approved_overtime_hours FLOAT DEFAULT 0"
            ))


def migrate_v19_payroll_columns():
    """Add V19 payroll attendance deduction field to existing payrolls."""
    from sqlalchemy import text
    with db.engine.begin() as conn:
        existing = _table_columns("payrolls")
        if "late_deduction" not in existing:
            conn.execute(text(
                "ALTER TABLE payrolls ADD COLUMN late_deduction FLOAT DEFAULT 0"
            ))


def migrate_v80_financial_link_columns():
    """Ensure V78/V79 financial-link columns exist on older production DBs."""
    from sqlalchemy import text
    with db.engine.begin() as conn:
        tables = set(inspect(db.engine).get_table_names())
        if "monthly_contributions" in tables:
            existing = _table_columns("monthly_contributions")
            if "source_payroll_id" not in existing:
                conn.execute(text("ALTER TABLE monthly_contributions ADD COLUMN source_payroll_id INTEGER"))
        if "expenses" in tables:
            existing = _table_columns("expenses")
            for col, definition in [
                ("source_type", "VARCHAR(40)"),
                ("source_key", "VARCHAR(100)"),
                ("payment_reference", "VARCHAR(150)"),
            ]:
                if col not in existing:
                    conn.execute(text(f"ALTER TABLE expenses ADD COLUMN {col} {definition}"))


def migrate_v28_contribution_columns():
    """Add V28 statutory-basis fields to existing monthly contribution records."""
    from sqlalchemy import text
    with db.engine.begin() as conn:
        columns = [
            ("sss_remuneration", "FLOAT DEFAULT 0"),
            ("sss_msc", "FLOAT DEFAULT 0"),
            ("philhealth_basic_salary", "FLOAT DEFAULT 0"),
            ("pagibig_monthly_compensation", "FLOAT DEFAULT 0"),
        ]
        existing = _table_columns("monthly_contributions")
        for col, definition in columns:
            if col not in existing:
                conn.execute(text(f"ALTER TABLE monthly_contributions ADD COLUMN {col} {definition}"))


def normalize_attendance_statuses():
    """Repair legacy attendance records with stale Present/Invalid Out status."""
    changed = False
    for record in Attendance.query.all():
        old = (record.status, record.total_hours, record.late_hours, record.undertime_hours, record.overtime_hours)
        calculate_attendance(record)
        new = (record.status, record.total_hours, record.late_hours, record.undertime_hours, record.overtime_hours)
        if old != new:
            changed = True
    if changed:
        db.session.commit()

def migrate_v42_columns():
    """Create Calendar/holiday table and holiday payroll column for existing V41 databases."""
    db.create_all()
    inspector = db.inspect(db.engine)
    tables = inspector.get_table_names()
    if "payrolls" in tables:
        cols = {c["name"] for c in inspector.get_columns("payrolls")}
        if "holiday_pay" not in cols:
            db.session.execute(db.text("ALTER TABLE payrolls ADD COLUMN holiday_pay FLOAT DEFAULT 0"))
    db.session.commit()


def migrate_v46_leave_columns():
    """Add V46 leave entitlement, usage and conversion fields."""
    inspector = db.inspect(db.engine)
    with db.engine.begin() as conn:
        if "employees" in inspector.get_table_names():
            cols = _table_columns("employees")
            additions = {
                "regularization_date": "DATE",
                "vl_entitlement": "FLOAT DEFAULT 0",
                "sl_entitlement": "FLOAT DEFAULT 0",
            }
            for name, ddl in additions.items():
                if name not in cols:
                    conn.execute(db.text(f"ALTER TABLE employees ADD COLUMN {name} {ddl}"))
        if "payrolls" in inspector.get_table_names():
            cols = _table_columns("payrolls")
            additions = {
                "vl_days": "INTEGER DEFAULT 0", "sl_days": "INTEGER DEFAULT 0",
                "vl_pay": "FLOAT DEFAULT 0", "sl_pay": "FLOAT DEFAULT 0",
                "vl_conversion": "FLOAT DEFAULT 0", "sl_conversion": "FLOAT DEFAULT 0",
            }
            for name, ddl in additions.items():
                if name not in cols:
                    conn.execute(db.text(f"ALTER TABLE payrolls ADD COLUMN {name} {ddl}"))

def migrate_v45_columns():
    """Add V45 payroll attendance/premium summary columns to existing databases."""
    inspector = db.inspect(db.engine)
    if "payrolls" not in inspector.get_table_names():
        return
    cols = {c["name"] for c in inspector.get_columns("payrolls")}
    additions = {
        "rest_day_premium": "FLOAT DEFAULT 0",
        "regular_ot_hours": "FLOAT DEFAULT 0",
        "regular_ot_pay": "FLOAT DEFAULT 0",
        "holiday_ot_hours": "FLOAT DEFAULT 0",
        "holiday_ot_pay": "FLOAT DEFAULT 0",
        "rest_day_ot_hours": "FLOAT DEFAULT 0",
        "rest_day_ot_pay": "FLOAT DEFAULT 0",
        "absent_days": "INTEGER DEFAULT 0",
        "leave_with_pay_days": "INTEGER DEFAULT 0",
        "leave_without_pay_days": "INTEGER DEFAULT 0",
    }
    with db.engine.begin() as conn:
        for name, ddl in additions.items():
            if name not in cols:
                conn.execute(db.text(f"ALTER TABLE payrolls ADD COLUMN {name} {ddl}"))


def migrate_v53_employee_profile_columns():
    """Add extended employee 201-file/profile fields to existing databases."""
    inspector = db.inspect(db.engine)
    if "employees" not in inspector.get_table_names():
        return
    additions = {
        "nickname": "VARCHAR(100)", "date_of_birth": "VARCHAR(30)", "place_of_birth": "VARCHAR(200)",
        "sex": "VARCHAR(30)", "civil_status": "VARCHAR(50)", "nationality": "VARCHAR(100)",
        "religion": "VARCHAR(100)", "blood_type": "VARCHAR(10)", "height": "VARCHAR(30)", "weight": "VARCHAR(30)",
        "permanent_address": "TEXT", "secondary_phone": "VARCHAR(50)", "company_email": "VARCHAR(150)",
        "social_media": "VARCHAR(255)", "professional_summary": "TEXT", "reporting_to": "VARCHAR(150)",
        "national_id": "VARCHAR(80)", "voters_id": "VARCHAR(80)", "drivers_license": "VARCHAR(80)",
        "passport_number": "VARCHAR(80)", "passport_expiry": "VARCHAR(30)",
        "work_location": "VARCHAR(255)", "employment_type": "VARCHAR(50)", "schedule_type": "VARCHAR(50)",
        "schedule_details": "VARCHAR(255)", "education_json": "TEXT", "work_experience_json": "TEXT",
        "skills_json": "TEXT", "family_json": "TEXT", "references_json": "TEXT", "documents_json": "TEXT",
        "duties": "TEXT",
    }
    with db.engine.begin() as conn:
        cols = _table_columns("employees")
        for name, ddl in additions.items():
            if name not in cols:
                conn.execute(db.text(f"ALTER TABLE employees ADD COLUMN {name} {ddl}"))


def migrate_v70_devices():
    """Create the device access-control table for existing EMS databases."""
    # db.create_all() creates the table for new deployments. For existing
    # SQLite/PostgreSQL databases this helper is intentionally a no-op because
    # the Device model has no changes to existing tables; db.create_all()
    # above/below creates the new table when absent.
    return None


def migrate_v47_leave_requests():
    """Ensure the employee leave-request workflow table exists for V47."""
    # db.create_all() creates the table from the LeaveRequest model. This helper
    # is intentionally kept for compatibility with existing V47 databases.
    return None

def seed_2026_holidays():
    """Seed nationwide 2026 holidays; Admin can edit/delete them in Calendar."""
    rows = [
        ("2026-01-01", "New Year's Day", "Regular Holiday"),
        ("2026-02-17", "Chinese New Year", "Special Non-Working Day"),
        ("2026-02-25", "EDSA People Power Revolution Anniversary", "Special Working Day"),
        ("2026-03-20", "Eid'l Fitr (Feast of Ramadhan)", "Regular Holiday"),
        ("2026-04-02", "Maundy Thursday", "Regular Holiday"),
        ("2026-04-03", "Good Friday", "Regular Holiday"),
        ("2026-04-04", "Black Saturday", "Special Non-Working Day"),
        ("2026-04-09", "Araw ng Kagitingan", "Regular Holiday"),
        ("2026-05-01", "Labor Day", "Regular Holiday"),
        ("2026-05-27", "Eid'l Adha (Feast of Sacrifice)", "Regular Holiday"),
        ("2026-06-12", "Independence Day", "Regular Holiday"),
        ("2026-08-21", "Ninoy Aquino Day", "Special Non-Working Day"),
        ("2026-08-31", "National Heroes Day", "Regular Holiday"),
        ("2026-11-01", "All Saints' Day", "Special Non-Working Day"),
        ("2026-11-02", "All Souls' Day", "Special Non-Working Day"),
        ("2026-11-30", "Bonifacio Day", "Regular Holiday"),
        ("2026-12-08", "Feast of the Immaculate Conception of Mary", "Special Non-Working Day"),
        ("2026-12-24", "Christmas Eve", "Special Non-Working Day"),
        ("2026-12-25", "Christmas Day", "Regular Holiday"),
        ("2026-12-30", "Rizal Day", "Regular Holiday"),
        ("2026-12-31", "Last Day of the Year", "Special Non-Working Day"),
    ]
    for ds, name, htype in rows:
        d = date.fromisoformat(ds)
        if not Holiday.query.filter_by(holiday_date=d).first():
            wm, um, om = holiday_defaults(htype)
            db.session.add(Holiday(holiday_date=d, name=name, holiday_type=htype,
                                   work_multiplier=wm, unworked_multiplier=um, ot_multiplier=om, active=True))
    db.session.commit()


@app.errorhandler(500)
def handle_internal_server_error(error):
    # Keep the browser response clean while writing the actual exception to
    # Render/Gunicorn logs for diagnosis instead of hiding the root cause.
    app.logger.exception("EMS Internal Server Error", exc_info=error)
    return "Internal Server Error", 500


def initialize_database():
    with app.app_context():
        db.create_all()
        migrate_v42_columns()
        migrate_v45_columns()
        migrate_v46_leave_columns()
        migrate_v47_leave_requests()
        migrate_v53_employee_profile_columns()
        migrate_v15_columns()
        migrate_v16_attendance_columns()
        migrate_v19_payroll_columns()
        migrate_v12_columns()
        migrate_v28_contribution_columns()
        migrate_v80_financial_link_columns()
        migrate_v59_email_columns()
        normalize_attendance_statuses()
        # Payroll no longer uses Draft. Existing payroll rows from older versions
        # are treated as Approved so they immediately follow the new workflow.
        db.session.query(Payroll).filter(Payroll.status == "Draft").update({"status": "Approved"}, synchronize_session=False)
        seed_2026_holidays()

        if not User.query.filter_by(username="admin").first():
            admin = User(
                username="admin",
                role="admin"
            )
            admin.set_password("admin123")
            db.session.add(admin)

        if not PayrollSettings.query.first():
            db.session.add(PayrollSettings())

        db.session.commit()


# Initialize the database whenever the application is imported by Gunicorn/Render.
# Gunicorn does not execute the __main__ block, so database tables and the default
# admin account must be created during application startup/import.
initialize_database()

if __name__ == "__main__":
    app.run(debug=True)

