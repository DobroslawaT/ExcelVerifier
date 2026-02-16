# Import/Export Documentation

## Overview

The Import/Export feature allows you to backup, restore, and migrate data between ExcelVerifier instances. Access it via the 📦 button in the top-right corner of the application.

## Features

### 1. Export All Data (📥 Eksportuj Wszystkie Dane)

Creates a complete backup of your application data in a ZIP archive.

**What's included:**
- ✅ Database (approved records, reporting data)
- ✅ All Excel files (approved and unapproved)
- ✅ All linked images (JPG, PNG, BMP)
- ✅ Company database (company_db.json)
- ✅ Application settings (settings.json)

**Use cases:**
- Regular backups
- Moving data to another computer
- Archiving old data
- Creating snapshots before major changes

**How to use:**
1. Click "📥 Eksportuj Wszystkie Dane"
2. Choose where to save the ZIP file
3. Wait for export to complete
4. Store the ZIP file safely

**Recommended:** Export weekly or before major changes.

---

### 2. Import from ZIP Archive (📦 Importuj z Archiwum ZIP)

Restore data from a previously exported ZIP file.

**Options:**
- **Replace mode** (default): Deletes all current data and replaces with archive contents
- **Merge mode** (checkbox): Combines archive data with existing data
  - Skips duplicate records
  - Adds new records
  - Doesn't overwrite existing files

**Use cases:**
- Restoring from backup
- Migrating from another computer
- Combining data from multiple sources (merge mode)

**How to use:**
1. ✅ **Optional:** Check "Połącz z istniejącymi danymi" for merge mode
2. Click "📦 Importuj z Archiwum ZIP"
3. Select the ZIP file to import
4. Confirm the operation
5. Wait for import to complete

**⚠️ Warning:** Replace mode will delete all current data! Create a backup first.

---

### 3. Import from Excel (📊 Importuj z Excel)

Import approved records from an ApprovedRecords.xlsx file with **automatic file detection**.

**What's imported:**
- Reads file paths from Excel (column: filepath)
- **Automatically finds and copies Excel files** from those paths
- **Automatically finds and copies associated images** (same name, different extension)
- Adds all records to database
- Skips duplicates automatically

**Use cases:**
- Migrating from old Excel-based system
- Importing records from another instance
- Batch importing files listed in an Excel manifest

**How to use:**
1. Prepare Excel file with columns: Date, Company, Filename, Filepath
2. Click "📊 Importuj z Excel"
3. Select your Excel file (e.g., ApprovedRecords.xlsx)
4. Confirm the automatic import process
5. Wait for completion
6. Review summary (files imported, images copied, errors)

**Excel format:**
```
| Date       | Company    | Filename          | Filepath                          |
|------------|------------|-------------------|-----------------------------------|
| 2026-01-15 | FIRMA_A    | report1.xlsx      | C:\OldData\FIRMA_A\report1.xlsx   |
| 2026-01-20 | FIRMA_B    | report2.xlsx      | D:\Archive\FIRMA_B\report2.xlsx   |
```

**What happens automatically:**
1. System reads each filepath from Excel
2. Copies Excel file to: `Reports/Zatwierdzone/{Company}/{Filename}`
3. Looks for image with same name: `report1.jpg`, `report1.png`, etc.
4. Copies image to the same destination folder
5. Adds record to database with new filepath

**Note:** Source files must exist at the paths specified in Excel. The system will report any missing files.

---

### 4. Import Folder (📁 Importuj Folder)

Batch import a whole folder of Excel files with their linked images.

**What's imported:**
- All Excel files (*.xlsx) in folder and subfolders
- Linked images (same name as Excel file)
- Automatically reads date and company from Excel cells (D1, B1)

**Options:**
- **Jako zatwierdzone** (default): Imports files as approved reports
  - Adds to database
  - Organizes by company folders
- **Jako niezatwierdzone**: Imports files as unapproved reports
  - Copies to unapproved folder
  - Ready for verification

**Use cases:**
- Migrating from another system
- Batch importing scanned invoices
- Moving historical data into the app
- Importing data from external sources

**How to use:**
1. Organize your files in a folder (Excel + images with matching names)
2. Select import mode (approved/unapproved)
3. Click "📁 Importuj Folder"
4. Select the folder to import
5. Wait for batch import to complete
6. Review the summary

**File naming convention:**
- Excel file: `2026-01-15_FIRMA_DOKUMENT.xlsx`
- Image file: `2026-01-15_FIRMA_DOKUMENT.jpg` (or .png, .bmp)

**Note:** If date/company cannot be read from Excel, uses current date and "Nieznana Firma".

---

## Best Practices

### Regular Backups
- Export weekly using "📥 Eksportuj Wszystkie Dane"
- Store backups in multiple locations (external drive, cloud)
- Name backups with dates: `ExcelVerifier_Backup_20260212_143022.zip`

### Before Major Changes
- Always export before:
  - Updating the application
  - Deleting many records
  - Importing large batches
  - Restoring from old backups

### Migrating to New Computer
1. Export all data on old computer
2. Install ExcelVerifier on new computer
3. Import the ZIP archive (replace mode)
4. Verify data integrity

### Merging Data from Multiple Sources
1. Export data from each source
2. Import first archive (replace mode)
3. Import additional archives (merge mode)
4. Check for duplicates

### Recovering from Excel-Based System
1. Export your ApprovedRecords.xlsx with full file paths
2. Use "📊 Importuj z Excel" - system will automatically copy all files
3. Verify imported data in "Zatwierdzone" dialog

### Importing Multiple Folders Efficiently  
Instead of manually selecting 100 folders:
1. Create an Excel file listing all file paths you want to import
2. Use "📊 Importuj z Excel" for automatic detection and copying
3. Much faster than individual folder selection!

---

## Troubleshooting

### Export fails
- Check disk space (exports can be large)
- Close any open Excel files
- Run as administrator if permissions issues

### Import fails
- Verify ZIP file is not corrupted
- Check available disk space
- Ensure no files are locked (close Excel)

### Merge creates duplicates
- Merge mode skips records with same filename
- If files have different names but same content, duplicates will occur
- Review imported data and delete duplicates manually

### Import folder skips files
- Check file format (must be .xlsx)
- Verify Excel files contain data in cells D1 (date) and B1 (company)
- Check error messages in import summary
- Review file permissions

### After import, some reports missing
- Check if filepaths in database still valid
- Use "📁 Importuj Folder" to re-import missing files
- Verify files copied to correct directories

---

## File Locations

After import, files are organized as:

```
ExcelVerifier/
├── excelverifier.db                    (Database)
├── company_db.json                     (Companies)
├── settings.json                       (Settings)
└── Reports/
    ├── Niezatwierdzone/               (Unapproved)
    │   ├── report1.xlsx
    │   └── report1.jpg
    └── Zatwierdzone/                  (Approved)
        ├── FIRMA A/
        │   ├── 2026-01-15_FIRMA_A.xlsx
        │   └── 2026-01-15_FIRMA_A.jpg
        └── FIRMA B/
            ├── 2026-01-20_FIRMA_B.xlsx
            └── 2026-01-20_FIRMA_B.jpg
```

---

## API for Developers

The `ImportExportHandler` class provides programmatic access:

```python
from core.import_export import ImportExportHandler

handler = ImportExportHandler()

# Export
success, message = handler.export_all_data("backup.zip")

# Import (merge mode)
success, message = handler.import_all_data("backup.zip", merge=True)

# Import from Excel
success, message = handler.import_from_excel_file("ApprovedRecords.xlsx")

# Import folder
success, message = handler.import_folder_batch("/path/to/folder", status="approved")
```

---

## Version History

- **v2.0** - Added Import/Export functionality
  - Full data backup/restore
  - Excel migration support
  - Batch folder import
  - Database-based storage
