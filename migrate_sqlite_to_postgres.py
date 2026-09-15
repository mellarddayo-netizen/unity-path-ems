"""One-time migration: copy the local SQLite EMS database into Render/PostgreSQL.

Usage (run locally from the EMS project folder):
  set DATABASE_URL=postgresql://...   # PowerShell: $env:DATABASE_URL="..."
  python migrate_sqlite_to_postgres.py

The script does NOT delete the source SQLite database. It creates the target schema,
copies rows in foreign-key order, and resets PostgreSQL identity sequences.
"""
import os
import sys
from pathlib import Path
from sqlalchemy import create_engine, MetaData, text

BASE_DIR = Path(__file__).resolve().parent
SOURCE_DB = Path(os.getenv("SOURCE_DB_PATH", str(BASE_DIR / "employee_system_v10.db")))
TARGET_URL = os.getenv("DATABASE_URL", "").strip()
if TARGET_URL.startswith("postgres://"):
    TARGET_URL = "postgresql://" + TARGET_URL[len("postgres://"):]

if not SOURCE_DB.exists():
    raise SystemExit(f"Source SQLite database not found: {SOURCE_DB}")
if not TARGET_URL.startswith(("postgresql://", "postgresql+psycopg://")):
    raise SystemExit("DATABASE_URL must be a PostgreSQL connection URL.")

source = create_engine(f"sqlite:///{SOURCE_DB}")
target = create_engine(TARGET_URL, pool_pre_ping=True)

src_meta = MetaData()
src_meta.reflect(bind=source)

tgt_meta = MetaData()
# Import models without importing app.py, so no startup/seed code runs here.
from extensions import db  # noqa: E402
from models import (  # noqa: F401,E402
    Employee, User, Attendance, Payroll, PayrollSettings, FinalPay,
    EmployeeLoan, LoanPayment, MonthlyContribution, Holiday, LeaveRequest,
)
tgt_meta = db.metadata
with target.begin() as conn:
    tgt_meta.create_all(bind=conn)

# Tables in dependency order. SQLAlchemy's sorted_tables handles most FK cases.
tables = [t for t in tgt_meta.sorted_tables if t.name in src_meta.tables]

# If there are FK cycles, process the remaining tables after the first pass.
ordered_names = []
for t in tables:
    ordered_names.append(t.name)
for name in src_meta.tables:
    if name in tgt_meta.tables and name not in ordered_names:
        ordered_names.append(name)

print(f"Source: {SOURCE_DB}")
print(f"Target: PostgreSQL")
print("Tables:", ", ".join(ordered_names))

with source.connect() as src_conn, target.begin() as tgt_conn:
    for name in ordered_names:
        src_table = src_meta.tables[name]
        tgt_table = tgt_meta.tables[name]
        rows = src_conn.execute(src_table.select()).mappings().all()
        if not rows:
            print(f"{name}: 0 rows")
            continue
        common = [c.name for c in tgt_table.columns if c.name in src_table.c]
        payload = [{k: row[k] for k in common} for row in rows]
        # Insert in manageable batches.
        for i in range(0, len(payload), 500):
            tgt_conn.execute(tgt_table.insert(), payload[i:i+500])
        print(f"{name}: {len(payload)} rows")

# Reset identity/serial sequences so new records don't collide with migrated IDs.
with target.begin() as conn:
    for table in tgt_meta.sorted_tables:
        if table.name not in src_meta.tables:
            continue
        pk = [c for c in table.primary_key.columns if c.name == "id"]
        if not pk:
            continue
        try:
            seq = conn.execute(
                text("SELECT pg_get_serial_sequence(:table_name, 'id')"),
                {"table_name": table.name},
            ).scalar()
            if not seq:
                continue
            max_id = conn.execute(text(f'SELECT MAX("id") FROM "{table.name}"')).scalar()
            if max_id is not None:
                conn.execute(text("SELECT setval(:seq, :value, true)"), {"seq": seq, "value": int(max_id)})
        except Exception as exc:
            print(f"Sequence warning for {table.name}: {exc}")

print("Migration complete. Source SQLite database was not modified.")
