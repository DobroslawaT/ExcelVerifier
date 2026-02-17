# Building ExcelVerifier Installer

This guide explains how to create a Windows installer for ExcelVerifier.

## Prerequisites

1. **Python 3.8+** with all dependencies installed
2. **Inno Setup 6.0+** - Download from https://jrsoftware.org/isdl.php
3. **Active virtual environment** with project dependencies

## Step-by-Step Build Process

### 1. Build the Executables

From the `ExcelVerifier` directory, run:

```powershell
# Activate virtual environment
.\.venv\Scripts\Activate.ps1

# Build executables
python build.py
```

This creates:
- `dist/ExcelVerifier/ExcelVerifier.exe` - Main application
- `dist/ExcelVerifier/configure_api.exe` - API key configuration utility
- All required DLLs and dependencies in `dist/ExcelVerifier/`

### 2. Create the Installer

With Inno Setup installed, run from the project root:

```powershell
# Option 1: Using Inno Setup compiler from command line
iscc installer.iss

# Option 2: Open installer.iss in Inno Setup GUI and click Build
```

The installer will be created at:
```
installer_output/ExcelVerifier_Setup.exe
```

### 3. Distribute

Share `ExcelVerifier_Setup.exe` with users. The installer will:
- Install the application to Program Files
- Prompt for Google Gemini API key during installation
- Create desktop shortcut (optional)
- Create Start Menu entries
- Configure API key with DPAPI encryption

## Installer Features

### API Key Configuration

During installation, users are prompted to enter their Google Gemini API key. This is optional - they can configure it later through the app's Settings dialog.

The API key is stored encrypted using Windows DPAPI (Data Protection API) for security.

### Installation Directory

Default: `C:\Program Files\ExcelVerifier\`

Users can change this during installation.

### What Gets Installed

- ExcelVerifier application (all .exe and .dll files)
- Application icon
- Reports directories (created on first run)
- SQLite database (created on first run)

### What's NOT Included

- Python runtime (embedded in .exe files)
- Virtual environment
- Source code
- Development tools

## Troubleshooting

### PyInstaller Issues

If build fails with "module not found" errors:

```powershell
pip install --upgrade pyinstaller
pip install --force-reinstall <missing-module>
python build.py
```

### Inno Setup Not Found

Add Inno Setup to PATH or use full path:

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

### Large .exe Size

The `--onedir` build creates a folder with many files but faster startup. For a single .exe file, change `build.py`:

```python
"--onefile",  # Instead of "--onedir"
```

Note: Single file is larger and slower to start.

## Updating the Application

To release an update:

1. Update version in `installer.iss`:
   ```
   #define MyAppVersion "1.1.0"
   ```

2. Rebuild:
   ```powershell
   python build.py
   iscc installer.iss
   ```

3. Distribute new `ExcelVerifier_Setup.exe`

## Testing the Installer

Before distributing:

1. Test installation on a clean Windows machine (or VM)
2. Verify API key configuration works
3. Test all features (image transform, report generation, etc.)
4. Test uninstallation
5. Verify no Python installation is required

## File Sizes (Approximate)

- `dist/ExcelVerifier/` folder: ~150-200 MB
- `ExcelVerifier_Setup.exe`: ~60-80 MB (compressed)
- After installation: ~150-200 MB

## Security Notes

- API keys are encrypted with Windows DPAPI (user-specific)
- Keys are stored in `settings.json` in the app directory
- Each Windows user has their own encrypted key
- Keys cannot be decrypted by other users or on other machines
