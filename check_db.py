from ExcelVerifier.core.database_handler import DatabaseHandler
from ExcelVerifier.config import DATABASE_FILE
import os

db = DatabaseHandler(DATABASE_FILE)
records = db.get_all_approved_records()
print(f'Found {len(records)} approved records')
for r in records:
    filename = r.get('filename')
    filepath = r.get('filepath')
    print(f"  - {filename}: {filepath}")
    if filepath:
        exists = os.path.exists(filepath)
        print(f"    File exists: {exists}")
