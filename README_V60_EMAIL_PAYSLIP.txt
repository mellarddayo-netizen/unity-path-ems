V60 - Email Payslip PDF + Payroll Delete UI

Changes:
- Email payslip PDF now follows the on-screen payslip layout more closely.
- Added diagonal SYSTEM GENERATED watermark to emailed PDF.
- Added acknowledgement/signature block to emailed PDF.
- Restored per-row Delete payroll action.
- Restored Select All for Delete and Delete Selected controls.
- Existing email send/status features remain.
- Added .gitignore to protect .env and local databases.

SMTP .env example:
MAIL_SERVER=smtp.gmail.com
MAIL_PORT=587
MAIL_USERNAME=admin@unitypathrecovery.com
MAIL_PASSWORD=YOUR_GOOGLE_APP_PASSWORD
MAIL_FROM=payroll@unitypathrecovery.com
MAIL_FROM_NAME=UNITY PATH RECOVERY AND COLLECTION SERVICES OPC
