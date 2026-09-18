EMS V68 – Persistent Profile Photo Fix

Fixes profile pictures not loading after moving profile uploads to Render persistent disk.

Changes:
- Adds /profile-photos/<filename> route that serves images from PROFILE_STORAGE_DIR.
- Updates employee list/profile/edit/self-profile/dashboard templates to use the persistent photo route.
- Keeps PROFILE_STORAGE_DIR=/var/data/profiles.

Important:
If an old employee photo was uploaded before persistent storage was configured and its file is no longer on the disk, re-upload that employee's photo once. New uploads will persist on the Render disk.
