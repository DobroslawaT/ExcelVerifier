# config.py
import base64
import json
import os
from pathlib import Path


def get_project_root():
    """Get the project root directory (parent of ExcelVerifier folder)."""
    # This file is in ExcelVerifier/ExcelVerifier/config.py
    # Project root is 2 levels up
    return Path(__file__).parent.parent.parent


def resolve_path(path_str):
    """Resolve a path string to absolute Path.
    
    If path is relative, resolves from project root.
    If path is absolute, uses it directly.
    """
    path = Path(path_str)
    if path.is_absolute():
        return path
    else:
        return (get_project_root() / path).resolve()


def load_settings():
    """Load settings from settings.json file."""
    project_root = get_project_root()
    settings_file = project_root / "settings.json"
    
    if settings_file.exists():
        try:
            with open(settings_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    
    # Default settings (relative paths)
    return {
        "reports_directory": "Reports/Niezatwierdzone",
        "approved_directory": "Reports/Zatwierdzone",
        "transform_directory": "Reports"
    }


def save_settings(settings):
    """Save settings to settings.json file."""
    settings_file = get_project_root() / "settings.json"
    try:
        with open(settings_file, "w", encoding="utf-8") as file_handle:
            json.dump(settings, file_handle, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Warning: Could not save settings: {e}")
        return False


def _encrypt_dpapi(value):
    """Encrypt a string with Windows DPAPI and return base64 text."""
    try:
        import win32crypt
    except Exception as e:
        raise RuntimeError("DPAPI unavailable (win32crypt not installed)") from e

    data = value.encode("utf-8")
    encrypted = win32crypt.CryptProtectData(data, None, None, None, None, 0)
    return base64.b64encode(encrypted).decode("ascii")


def _decrypt_dpapi(encoded):
    """Decrypt a base64 DPAPI string and return plaintext."""
    try:
        import win32crypt
    except Exception as e:
        raise RuntimeError("DPAPI unavailable (win32crypt not installed)") from e

    data = base64.b64decode(encoded.encode("ascii"))
    decrypted = win32crypt.CryptUnprotectData(data, None, None, None, 0)[1]
    return decrypted.decode("utf-8")


def get_gemini_api_key():
    """Get Gemini API key from environment or encrypted settings."""
    env_key = os.environ.get("GEMINI_API_KEY")
    if env_key:
        return env_key

    encrypted = _settings.get("gemini_api_key_encrypted")
    if not encrypted:
        return None

    try:
        return _decrypt_dpapi(encrypted)
    except Exception as e:
        print(f"Warning: Could not decrypt GEMINI API key: {e}")
        return None


def set_gemini_api_key(api_key):
    """Encrypt and store Gemini API key in settings.json."""
    if not api_key:
        return False

    try:
        encrypted = _encrypt_dpapi(api_key)
    except Exception as e:
        print(f"Warning: Could not encrypt GEMINI API key: {e}")
        return False

    settings = load_settings()
    settings["gemini_api_key_encrypted"] = encrypted
    if not save_settings(settings):
        return False

    global _settings
    _settings = settings
    return True


def ensure_directories(*paths):
    """Ensure directories exist, create if missing."""
    for path in paths:
        try:
            path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"Warning: Could not create directory {path}: {e}")


# Load settings
_settings = load_settings()

# Resolve directory paths as Path objects
_REPORTS_ROOT_PATH = resolve_path(_settings.get("reports_directory", "Reports/Niezatwierdzone"))
_APPROVED_DIRECTORY_PATH = resolve_path(_settings.get("approved_directory", "Reports/Zatwierdzone"))
_TRANSFORM_DIRECTORY_PATH = resolve_path(_settings.get("transform_directory", "Reports"))

# Auto-create directories
ensure_directories(_REPORTS_ROOT_PATH, _APPROVED_DIRECTORY_PATH, _TRANSFORM_DIRECTORY_PATH)

# Export as strings for backward compatibility
REPORTS_ROOT = str(_REPORTS_ROOT_PATH)
APPROVED_DIRECTORY = str(_APPROVED_DIRECTORY_PATH)
TRANSFORM_DIRECTORY = str(_TRANSFORM_DIRECTORY_PATH)

# File paths
DEFAULT_IMAGE = None  # Set by user as needed
APPROVED_FILE = str(_APPROVED_DIRECTORY_PATH / "ApprovedRecords.xlsx")
REPORTING_DATA_FILE = "reportingData.xlsx"
COMPANY_DB_FILE = str(get_project_root() / "company_db.json")
DATABASE_FILE = str(get_project_root() / "excelverifier.db")

# Ensure parent directories for files exist
ensure_directories(get_project_root())

# Export Path versions for modern code (optional)
REPORTS_ROOT_PATH = _REPORTS_ROOT_PATH
APPROVED_DIRECTORY_PATH = _APPROVED_DIRECTORY_PATH
TRANSFORM_DIRECTORY_PATH = _TRANSFORM_DIRECTORY_PATH
APPROVED_FILE_PATH = _APPROVED_DIRECTORY_PATH / "ApprovedRecords.xlsx"
COMPANY_DB_FILE_PATH = get_project_root() / "company_db.json"
DATABASE_FILE_PATH = get_project_root() / "excelverifier.db"