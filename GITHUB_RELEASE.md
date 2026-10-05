# Noor-ul-Hifz 2.0

Offline student daily monitoring and reporting app for Hifz & Qur'an institutes.

## What's included
- Fixed-range Excel daily report import
- Manual student entry
- Dashboard
- Collect Reports
- All Students
- Local SQLite persistence
- Same-day replacement/upsert for corrections
- Future-date protection
- Offline/private workflow
- Windows PyInstaller build script

## Important data location
When running as a packaged Windows EXE, data is stored at:
`%LOCALAPPDATA%\Noor-ul-Hifz\student_logs.db`

Do not upload the database if it contains real student information.

## Build
Run `BUILD_WINDOWS.bat` on Windows, or use:

`py -m pip install --upgrade pyinstaller openpyxl`

`pyinstaller --noconfirm --clean --windowed --onefile --name "Noor-ul-Hifz" --icon "icon.ico" main.py`
