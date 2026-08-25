@echo off
cd /d "%~dp0"

set "FOLDER_PATH_gme_compliance=C:\Users\USERNAME\OneDrive - Cedars-Sinai\GME"

echo Running GME Weekly Compliance generator...
python work_hours_compliance_generator.py

echo.
echo Process complete.
echo Check the shared GME folder for weekly_compliance_email_list.xlsx.
pause