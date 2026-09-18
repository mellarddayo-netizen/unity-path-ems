EMS V69 – PROFILE PHOTO ROUTE FIX

This version fixes the Render 500 error:
BuildError: Could not build url for endpoint profile_photo

Changes:
- Explicitly registers the Flask endpoint as profile_photo.
- Employee photo templates use the stable /profile-photos/<filename> path.
- Profile photos continue to use PROFILE_STORAGE_DIR (Render: /var/data/profiles).
- Missing photos return 404 instead of breaking the Employees page.
- PostgreSQL driver remains psycopg2-binary==2.9.10.

IMPORTANT DEPLOYMENT:
Upload/replace BOTH app.py and the templates folder from this package. Do not upload only templates.
Keep the existing Render Environment Variables, DATABASE_URL, PROFILE_STORAGE_DIR, disk, and PostgreSQL unchanged.
