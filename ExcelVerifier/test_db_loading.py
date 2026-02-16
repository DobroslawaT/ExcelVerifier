#!/usr/bin/env python3
"""Direct test of database approved records loading"""

import sys
import os

# Path fix
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

try:
    print("=" * 70)
    print("DATABASE APPROVED RECORDS LOADING TEST")
    print("=" * 70)
    
    print("\n[1] Importing DatabaseHandler...")
    from ExcelVerifier.core.database_handler import DatabaseHandler
    from ExcelVerifier.config import DATABASE_FILE
    
    print(f"[2] DATABASE_FILE = {DATABASE_FILE}")
    print(f"[3] Does DATABASE_FILE exist? {os.path.exists(DATABASE_FILE)}")
    
    print("\n[4] Creating DatabaseHandler...")
    db = DatabaseHandler(DATABASE_FILE)
    
    print("[5] Calling get_all_approved_records()...")
    records = db.get_all_approved_records()
    
    print(f"\n[6] Result: {len(records)} records returned")
    
    if records:
        print(f"\n[7] Record details:")
        for i, record in enumerate(records):
            print(f"\n    Record {i}:")
            for key, value in record.items():
                print(f"      {key}: {value} ({type(value).__name__})")
    else:
        print("\n[7] No records returned!")
    
    print("\n" + "=" * 70)
    print("✓ TEST COMPLETE")
    print("=" * 70)

except Exception as e:
    print(f"\n✗ ERROR: {e}")
    import traceback
    traceback.print_exc()
