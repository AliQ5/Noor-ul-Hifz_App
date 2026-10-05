# نور الحفظ — Noor-ul-Hifz V3

Qur'an Memorization Student Monitoring & Daily Reporting System.

## V3 fixes

### 1. Persistent database — critical fix
The previous one-file Windows EXE stored SQLite beside `__file__`. With PyInstaller `--onefile`, that location is the temporary extraction directory, which is removed when the application closes. That made imported data appear to save, then disappear after reopening.

V3 stores the database in a persistent Windows user-data directory when packaged:

```text
%LOCALAPPDATA%\Noor-ul-Hifz\student_logs.db
```

The database is created automatically and survives application restarts and EXE replacement.

When running from Python source, the database remains in the project's `data/` folder.

### 2. Manual entry
A new **Manual Entry** section allows staff to record a student without Excel:

- Date
- Student ID
- Student name
- غیر حاضری
- تاخیر
- سبق کا ناغہ
- سبقی کا ناغہ
- منزل کا ناغہ

Existing students can be loaded by ID.

### 3. Corrected same-day updates
Same student + same date is an upsert. A corrected import/manual entry now **replaces** the five daily flags instead of using `MAX()` and accidentally preserving old flags.

### 4. Website-aligned visual direction
The desktop UI now follows the Noor-ul-Hifz website's calmer visual language: warm/off-white surfaces, dark ink typography, restrained blue accent, thin borders, generous spacing, and simple data-focused cards.

Website: https://noorulhifz.netlify.app/

## Excel import

The importer still uses the fixed ranges from the institute's daily workbook:

| Category | ID ranges | Name ranges |
|---|---|---|
| غیر حاضری | B4:B7, F4:F7 | C4:C7, G4:G7 |
| تاخیر | B11:B22, F11:F22 | C11:C22, G11:G22 |
| سبق کا ناغہ | B26:B39, F26:F39 | C26:C39, G26:G39 |
| سبقی کا ناغہ | B43:B54, F43:F54 | C43:C54, G43:G54 |
| منزل کا ناغہ | B58:B69, F58:F69 | C58:C69, G58:G69 |

Only the `یومیہ رپورٹ` sheet is read.

## Run from source

```bat
python -m pip install -r requirements.txt
python main.py
```

## Build the Windows EXE

```bat
pyinstaller --noconfirm --clean --windowed --onefile --name "Noor-ul-Hifz" --icon "icon.ico" main.py
```

The EXE will be in:

```text
dist\Noor-ul-Hifz.exe
```

## Important test before deployment

1. Build the EXE.
2. Open it.
3. Add one record using Manual Entry.
4. Close the EXE completely.
5. Open it again.
6. Open **All Students** and **Collect Report**.
7. Confirm the record is still present.
8. Import an Excel report.
9. Close and reopen again.
10. Confirm the imported records remain.

Do not test persistence by looking only at the success message. The real test is closing the application and opening it again.

## Data backup

Back up this file on the Windows machine:

```text
%LOCALAPPDATA%\Noor-ul-Hifz\student_logs.db
```

Do not publish a real institute database in GitHub releases.

## Author

**ALI QURESHI**

Phone: `0300-2917500`  
Email: `aliaqureshi75@gmail.com`
