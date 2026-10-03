# نور الحفظ --- Noor-ul-Hifz

> **Qur'an Memorization Student Monitoring & Daily Reporting System**

**Noor-ul-Hifz (نور الحفظ)** is a lightweight Windows desktop
application designed to help Qur'an memorization institutes manage daily
student attendance, lateness, and lesson-related records in an
organized, offline-first system.

It replaces scattered daily Excel sheets and manual monthly counting
with a structured local database, Excel importing, reporting, and
student summaries.

------------------------------------------------------------------------

## ✨ Overview

Noor-ul-Hifz is built for institutes where teachers or office staff need
to record daily student status such as:

-   **غیر حاضری** --- Absent
-   **تاخیر** --- Late
-   **سبق کا ناغہ** --- Sabaq not completed
-   **سبقی کا ناغہ** --- Sabqi not completed
-   **منزل کا ناغہ** --- Manzil not completed

The application stores these records locally and allows the office to
review student history and generate reports without requiring an
internet connection.

------------------------------------------------------------------------

## 🎯 Main Goals

-   Reduce repetitive manual data entry.
-   Convert existing daily Excel reports into structured records.
-   Keep a complete history for every student.
-   Make monthly/overall counting fast and reliable.
-   Keep student data stored locally.
-   Provide simple Excel export for office use.
-   Make the workflow easy enough for everyday institute staff.

------------------------------------------------------------------------

## 🖥️ Core Features

### 1. Daily Excel Import

Import the institute's existing daily Excel report directly into
Noor-ul-Hifz.

The importer reads the predefined sections from the daily report:

  Category       Excel ID Range         Excel Name Range
  -------------- ---------------------- ----------------------
  غیر حاضری      `B4:B7`, `F4:F7`       `C4:C7`, `G4:G7`
  تاخیر          `B11:B22`, `F11:F22`   `C11:C22`, `G11:G22`
  سبق کا ناغہ    `B26:B39`, `F26:F39`   `C26:C39`, `G26:G39`
  سبقی کا ناغہ   `B43:B54`, `F43:F54`   `C43:C54`, `G43:G54`
  منزل کا ناغہ   `B58:B69`, `F58:F69`   `C58:C69`, `G58:G69`

Only these ranges are used for the daily import. Other worksheets are
ignored.

------------------------------------------------------------------------

### 2. Import Review

Before saving an imported report, the application displays a review
screen containing:

-   Selected filename
-   Total records detected
-   Absent count
-   Late count
-   Sabaq count
-   Sabqi count
-   Manzil count
-   Import date

This gives the user an opportunity to verify the data before committing
it.

------------------------------------------------------------------------

### 3. Manual Date Selection

The import date can be selected manually.

By default, Noor-ul-Hifz uses the current date.

Rules:

-   Date format: `YYYY-MM-DD`
-   Past dates are allowed.
-   The current date is allowed.
-   Future dates are rejected.

This is useful when the office enters a report after the actual school
day.

------------------------------------------------------------------------

### 4. Duplicate-Safe Daily Records

Importing the same date again does not intentionally create duplicate
daily records.

Existing records for that student/date combination are updated according
to the latest imported information.

This makes it possible to correct a daily report and import it again.

------------------------------------------------------------------------

### 5. All Students

The **All Students** section provides a complete student overview.

Each student can be viewed with totals for:

  Column         Meaning
  -------------- ---------------------
  ID             Student ID
  Name           Student name
  غیر حاضری      Total absent days
  تاخیر          Total late days
  سبق کا ناغہ    Total Sabaq misses
  سبقی کا ناغہ   Total Sabqi misses
  منزل کا ناغہ   Total Manzil misses

The section also displays the total number of students stored in the
system.

------------------------------------------------------------------------

### 6. Collect Reports

The reporting workflow is designed to answer questions such as:

> How many times did this student miss Sabaq?

> How many late days does this student have?

> How many Manzil misses were recorded?

Student ID can be used as the primary identifier for retrieving stored
records.

------------------------------------------------------------------------

### 7. Excel Export

Student data and calculated totals can be exported to Excel for:

-   Office records
-   Printing
-   Monthly review
-   Administrative reporting
-   Further analysis

------------------------------------------------------------------------

### 8. Local / Offline Database

Noor-ul-Hifz uses **SQLite** for local storage.

The application does not require an online database or server for its
core functionality.

Benefits:

-   Fast
-   Lightweight
-   Offline
-   No server setup
-   No internet dependency
-   Simple backup process

------------------------------------------------------------------------

## 🔐 Privacy

Student records are intended to remain on the local computer.

Noor-ul-Hifz is designed as an offline-first application and does not
require student information to be uploaded to a cloud service for normal
operation.

> **Important:** Anyone who has access to the computer or database files
> may potentially access the stored records. Use normal Windows account
> security and backups to protect institute data.

------------------------------------------------------------------------

## 📁 Data Storage

The application uses a local SQLite database.

Typical project structure:

``` text
Noor-ul-Hifz/
├── Noor-ul-Hifz.exe
├── data/
│   └── hifz_logs.db
└── ...
```

The database contains the application's locally stored student and
daily-log information.

### Backup

For a portable/local deployment, regularly back up the database file:

``` text
data/hifz_logs.db
```

A simple backup can be made by copying the database file to another
secure location.

------------------------------------------------------------------------

## 📊 Daily Workflow

The intended office workflow is:

``` text
Teacher / Daily Report
        ↓
Existing Excel Sheet
        ↓
Select Excel File
        ↓
Review Imported Records
        ↓
Choose / Confirm Date
        ↓
Import & Save
        ↓
SQLite Database
        ↓
All Students / Collect Reports
        ↓
Excel Export
```

------------------------------------------------------------------------

## 📋 Supported Daily Categories

### غیر حاضری

Student was absent.

### تاخیر

Student arrived late.

### سبق کا ناغہ

Student did not complete the assigned Sabaq.

### سبقی کا ناغہ

Student did not complete the assigned Sabqi.

### منزل کا ناغہ

Student did not complete the assigned Manzil.

A single student can have multiple conditions recorded on the same day
when applicable.

------------------------------------------------------------------------

## 🆔 Student Identification

The **Student ID** is treated as the primary student identifier.

Names are imported from the corresponding Excel cells.

For reliable long-term reporting:

-   Keep student IDs consistent.
-   Avoid changing an existing student's ID.
-   Do not reuse an ID for another student.
-   Keep the daily Excel format consistent.

------------------------------------------------------------------------

## 📑 Excel Template Requirements

The daily workbook must contain the expected first-sheet layout.

The application currently uses:

``` text
Sheet: یومیہ رپورٹ
```

The relevant data is read from the fixed cell ranges documented above.

The other workbook sheets are not required for the import process.

### Important

If the institute changes the layout of the daily Excel report, the
importer ranges will need to be updated in the application source code.

------------------------------------------------------------------------

## 🛠️ Technology

Noor-ul-Hifz is designed as a lightweight Python desktop application.

Core technologies:

-   **Python**
-   **Tkinter** --- Desktop interface
-   **SQLite** --- Local database
-   **OpenPyXL** --- Excel reading/writing
-   **PyInstaller** --- Windows executable packaging

The application is designed to work without a web server.

------------------------------------------------------------------------

## 💻 Development Setup

For development, install Python and the required packages.

``` bash
python -m pip install --upgrade pip
python -m pip install openpyxl
```

Run the application from the project source:

``` bash
python main.py
```

If the main source file has a different name, use that filename instead.

------------------------------------------------------------------------

## 📦 Building the Windows EXE

Install PyInstaller:

``` bash
python -m pip install pyinstaller
```

Build a single Windows executable:

``` bash
pyinstaller --noconfirm --clean --windowed --onefile --name "Noor-ul-Hifz" main.py
```

The executable will be created at:

``` text
dist\Noor-ul-Hifz.exe
```

### With an Application Icon

If `icon.ico` is available:

``` bash
pyinstaller --noconfirm --clean --windowed --onefile --name "Noor-ul-Hifz" --icon "icon.ico" main.py
```

------------------------------------------------------------------------

## 🧰 Recommended Distribution

For regular institute use, the preferred release structure is:

``` text
Noor-ul-Hifz/
├── Noor-ul-Hifz.exe
├── data/
│   └── hifz_logs.db
└── README.md
```

For a more polished public release, the project can also be distributed
through a GitHub Release and packaged with a Windows installer.

------------------------------------------------------------------------

## 🔄 Updating the Application

When updating the application:

1.  Close Noor-ul-Hifz.
2.  Back up the database.
3.  Replace the application executable.
4.  Keep the existing database.
5.  Start the new version.
6.  Verify the student/report data.

**Never delete the database when updating unless you intentionally want
to start with a fresh system.**

------------------------------------------------------------------------

## ⚠️ Important Data-Safety Notes

Before making major changes:

-   Back up `hifz_logs.db`.
-   Keep a copy of important Excel reports.
-   Do not reuse student IDs.
-   Do not modify the Excel template without updating the importer.
-   Test a new application version with a backup database first.

------------------------------------------------------------------------

## 🧪 Recommended Testing Before Production

Before deploying a new version to the institute, test:

-   [ ] Excel import
-   [ ] Import review
-   [ ] Manual date selection
-   [ ] Future-date rejection
-   [ ] Same-date re-import
-   [ ] Student totals
-   [ ] All Students
-   [ ] Collect Reports
-   [ ] Excel export
-   [ ] Database persistence
-   [ ] Application restart
-   [ ] Database backup/restore

------------------------------------------------------------------------

## 🎨 Product Identity

### Name

**نور الحفظ**

### English Name

**Noor-ul-Hifz**

### Meaning

**نور الحفظ** can be understood as **"Light of Memorization."**

### Visual Direction

The application uses a clean professional visual identity combining:

-   Blue accent colors
-   White and soft gray surfaces
-   Modern desktop UI
-   Minimal Islamic visual language
-   Arabic branding
-   Clear data-focused layouts

------------------------------------------------------------------------

## 👨‍💻 Author

### ALI QURESHI

**Software Developer**

For inquiries, support, customization, or deployment:

**Phone:** `0300-2917500`

**Email:** `aliaqureshi75@gmail.com`

------------------------------------------------------------------------

## 📄 License

Copyright © ALI QURESHI.

Unless a separate license is provided with a specific release, the
source code, branding, application design, and associated materials
should not be redistributed, modified, or commercially reused without
permission from the author.

------------------------------------------------------------------------

## 📬 Support

For application support or reporting an issue, contact:

**ALI QURESHI**

📞 `0300-2917500`\
✉️ `aliaqureshi75@gmail.com`

When reporting a problem, include:

1.  Noor-ul-Hifz version
2.  Windows version
3.  What you were trying to do
4.  The exact error message, if any
5.  A copy of the relevant Excel file when appropriate

------------------------------------------------------------------------

## 🌙 Noor-ul-Hifz

**Simple records. Clear reports. Better student management.**

**نور الحفظ --- Qur'an Memorization Student Monitoring System**
