# نور الحفظ — Noor-ul-Hifz V4

Qur'an Memorization Student Monitoring & Daily Reporting System.

## V4 Major Update

### 1. Multi-Dimensional Record Filtering (Collect Student Report)
When reviewing any student's record history, staff can now instantly filter records by:
- **Specific Date**: Inspect a single day (`YYYY-MM-DD`) with one-click "Today" shortcut.
- **Day of Week**: Look up records for specific days across the entire term (e.g. Every Friday / پیر، منگل، بدھ، جمعرات، جمعہ، ہفتہ، اتوار).
- **Week Number**: Filter by week of the year (Week 1–53 + Year).
- **Month**: Review any calendar month (e.g. October 2026).
- **Year**: View complete annual logs.
- **Dynamic Summaries & Filtered Totals**: Shows both filtered totals and lifetime totals side-by-side.
- **Filtered Excel Export**: Exports styled, filtered spreadsheets with filter parameters preserved in the report header.

### 2. Appearance & Customisation (App Settings)
- **Theme Selection**: Toggle between crisp Light Mode and immersive Dark Mode with instant live application.
- **Accent Color Palettes**: Choose from 8 curated presets (*Classic Blue, Ocean Teal, Emerald Green, Royal Violet, Crimson Rose, Warm Amber, Slate Indigo, Obsidian*) or enter any custom Hex color.
- **Font Scaling**: Adjust interface font sizes across Small, Normal, and Large.
- Settings are saved automatically and restored on subsequent launches.

### 3. V4 Performance & Concurrency Engine
- Enabled SQLite **WAL (Write-Ahead Logging)** mode for lightning-fast reads and writes.
- Multi-column index caching on `(student_id, log_date)` and `log_date`.
- Zebra-striped treeviews and active navigation pill indicators for a polished, modern feel.

### 4. Excel Import & Manual Entry
- Full support for institute Excel daily report workbooks (`یومیہ رپورٹ` sheet).
- Flexible manual student entry with real-time condition checkboxes.
- Overwrite protection, future date prevention, and automated safety backups.

---

## Run from source

```bat
python -m pip install -r requirements.txt
python main.py
```

## Build the Windows EXE

```bat
BUILD_WINDOWS.bat
```

or manually:

```bat
pyinstaller --noconfirm --clean --windowed --onefile --name "Noor-ul-Hifz" --icon "icon.ico" main.py
```

The compiled standalone executable will be generated at:
```text
dist\Noor-ul-Hifz.exe
```

## Inno Setup Windows Installer
To build the setup installer, compile `installer.iss` using Inno Setup Compiler. The setup wizard executable is saved to:
```text
release\Noor-ul-Hifz-Setup-4.0.exe
```

## Data Location & Backup
When packaged as an EXE, persistent data is stored in the local Windows user profile:
```text
%LOCALAPPDATA%\Noor-ul-Hifz\student_logs.db
```

Use the built-in "💾 Backup database" button in the sidebar or copy the database file directly for offline backups.

---

## Author
**ALI QURESHI**  
Phone: `0300-2917500`  
Email: `aliaqureshi75@gmail.com`  
Website: https://noorulhifz.netlify.app/
