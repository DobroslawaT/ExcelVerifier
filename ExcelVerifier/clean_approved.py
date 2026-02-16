#!/usr/bin/env python3
"""Script to clean approved records database and delete approved files."""

import os
import sys
import shutil
from pathlib import Path

# Add the project to path
sys.path.insert(0, os.path.dirname(__file__))

from ExcelVerifier.core.database_handler import DatabaseHandler
from ExcelVerifier.config import DATABASE_FILE

def main():
    print("\n" + "="*80)
    print("🗑️  CLEANING APPROVED RECORDS - DATABASE & FILES")
    print("="*80 + "\n")
    
    confirm = input("⚠️  This will delete ALL approved records and their files. Continue? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("❌ Cancelled.\n")
        return 0
    
    try:
        # 1. Clear database
        print("\n🗑️  Clearing approved_records table...")
        
        try:
            import sqlite3
            conn = sqlite3.connect(DATABASE_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM approved_records")
            conn.commit()
            conn.close()
            print("✅ Database cleared\n")
        except Exception as e:
            print(f"❌ Error clearing database: {e}\n")
            return 1
        
        # 2. Delete approved files
        print("🗑️  Deleting approved files...")
        reports_dir = Path("Reports/Zatwierdzone")
        
        if reports_dir.exists():
            try:
                shutil.rmtree(reports_dir)
                print(f"✅ Deleted: {reports_dir}")
            except Exception as e:
                print(f"❌ Error deleting {reports_dir}: {e}\n")
                return 1
        else:
            print(f"ℹ️  Directory not found: {reports_dir}")
        
        # Recreate the directory
        reports_dir.mkdir(parents=True, exist_ok=True)
        print(f"✅ Recreated empty directory: {reports_dir}\n")
        
        print("="*80)
        print("✅ APP IS NOW CLEAN - READY FOR TESTING")
        print("="*80 + "\n")
        return 0
        
    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}\n")
        return 1

if __name__ == "__main__":
    sys.exit(main())
