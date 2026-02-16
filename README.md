# ExcelVerifier

Excel report verification and transformation application with AI-powered image processing.

## Features

- Transform images to Excel using Google Gemini AI
- Verify and approve delivery reports
- Generate monthly rotation reports (butlodni)
- Company database management
- Data import/export with backup functionality
- Secure DPAPI encryption for API keys (Windows)

## Requirements

- Python 3.8+
- Windows OS (for DPAPI encryption)
- Google Gemini API key

## Installation

1. Clone the repository:
```bash
git clone git@github.com:DobroslawaT/ExcelVerifier.git
cd ExcelVerifier
```

2. Create and activate virtual environment:
```bash
python -m venv .venv
.venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r ExcelVerifier\ExcelVerifier\requirements.txt
```

4. Set up your Gemini API key (one of):
   - Set environment variable: `$env:GEMINI_API_KEY = "your-key-here"`
   - Or configure through the app's Settings dialog (stores encrypted with DPAPI)

## Running the Application

```bash
python ExcelVerifier\ExcelVerifier\main.py
```

## Project Structure

```
ExcelVerifier/
├── ExcelVerifier/          # Main application package
│   ├── core/               # Core business logic
│   │   ├── database_handler.py
│   │   ├── excel_handler.py
│   │   ├── image_transformer.py
│   │   └── import_export.py
│   ├── ui/                 # PyQt5 user interface
│   │   ├── main_window.py
│   │   ├── VerificationPage.py
│   │   ├── TransformPicToExcelPage.py
│   │   └── GenerateReportPage.py
│   ├── config.py           # Configuration with DPAPI encryption
│   └── main.py             # Application entry point
├── Reports/                # Generated reports (not in git)
│   ├── Zatwierdzone/       # Approved reports
│   └── Niezatwierdzone/    # Pending reports
└── excelverifier.db        # SQLite database (auto-created)
```

## Database

The SQLite database (`excelverifier.db`) is created automatically on first run with tables for:
- Companies
- Products  
- Orders
- Order items
- Approved records

## Documentation

See Polish documentation files:
- [DOKUMENTACJA.md](ExcelVerifier/DOKUMENTACJA.md) - Main documentation
- [DATABASE_MIGRATION.md](ExcelVerifier/DATABASE_MIGRATION.md) - Database migration guide
- [IMPORT_EXPORT.md](ExcelVerifier/IMPORT_EXPORT.md) - Import/export functionality

## Building Executable

```bash
python ExcelVerifier\build.py
```

This creates a standalone `.exe` in the `dist/` folder using PyInstaller.
