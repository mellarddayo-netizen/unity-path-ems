V65 RENDER PRODUCTION PREPARATION

This version keeps local PyCharm + SQLite working, while supporting PostgreSQL when
DATABASE_URL is present (Render).

IMPORTANT:
- Do not commit .env or employee_system_v10.db to GitHub.
- Render Free filesystem is ephemeral. Use Render Postgres for relational data.
- Profile photos are configurable through PROFILE_STORAGE_DIR. On Render, point this
to a mounted persistent disk such as /var/data/profiles if you need photos to survive redeploys.

RENDER SETTINGS
Build command:
  pip install -r requirements.txt
Start command:
  gunicorn app:app

ENVIRONMENT VARIABLES
  DATABASE_URL=<Render Postgres Internal Database URL>
  MAIL_SERVER=smtp.gmail.com
  MAIL_PORT=587
  MAIL_USERNAME=admin@unitypathrecovery.com
  MAIL_PASSWORD=<Google App Password>
  MAIL_FROM=payroll@unitypathrecovery.com
  MAIL_FROM_NAME=UNITY PATH RECOVERY AND COLLECTION SERVICES OPC
  PROFILE_STORAGE_DIR=/var/data/profiles   # only when a persistent disk is attached

ONE-TIME DATA MIGRATION
1. Create the Render Postgres database.
2. Copy its External Database URL temporarily to your local terminal as DATABASE_URL.
3. Put your current employee_system_v10.db beside this script.
4. Run:
   python migrate_sqlite_to_postgres.py
5. Verify row counts before switching the live service to the new database.
6. Never delete the original SQLite database until the production data has been verified.

EMAIL
Keep the Google App Password in Render Environment Variables only. Never put it in GitHub.

RECOMMENDED PRODUCTION ARCHITECTURE FOR ~20 EMPLOYEES
- Render Web Service: 0.5 CPU / 512 MB paid plan is enough to start.
- Render Postgres: 0.1 CPU / 256 MB paid plan is enough to start.
- Persistent disk: 5 GB is already a large allowance for employee profile photos; disks are billed separately.
- Keep the existing unity-path-ems Render service; do not create a second EMS service.
- Later, add ems.unitypathrecovery.com as a custom domain if desired. Keep unitypathrecovery.com for the company website.
