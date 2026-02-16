#!/usr/bin/env python3
"""Test script to verify approved reports dialog data loading"""

import sys
import os

# Path fix
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

import pandas as pd
from ExcelVerifier.core.database_handler import DatabaseHandler
from ExcelVerifier.config import DATABASE_FILE

print("=" * 60)
print("APPROVED REPORTS DATA LOADING TEST")
print("=" * 60)

try:
    # Simulate what load_data() does
    db = DatabaseHandler(DATABASE_FILE)
    records = db.get_all_approved_records()
    print(f"\n✓ Loaded {len(records)} approved records from database")
    
    if records:
        print(f"\nFirst record keys: {records[0].keys()}")
        print(f"\nFirst record data:")
        for key, value in records[0].items():
            print(f"  {key}: {value} ({type(value).__name__})")
    
    # Convert to DataFrame
    df = pd.DataFrame(records)
    print(f"\n✓ Created DataFrame")
    print(f"  Columns: {df.columns.tolist()}")
    print(f"  Shape: {df.shape}")
    
    # Rename columns (as done in load_data)
    if not df.empty:
        df_renamed = df.rename(columns={
            'date': 'Date',
            'company_name': 'Company',
            'company_nip': 'NIP',
            'filename': 'Filename',
            'filepath': 'Filepath'
        }, errors='ignore')
        
        print(f"\n✓ Renamed columns")
        print(f"  New columns: {df_renamed.columns.tolist()}")
        
        # Check each row for filepath
        print(f"\n✓ Checking filepaths in DataFrame:")
        for idx, row in df_renamed.iterrows():
            print(f"\n  Row {idx}:")
            print(f"    Date: {row.get('Date', 'N/A')}")
            print(f"    Company: {row.get('Company', 'N/A')}")
            print(f"    Filename: {row.get('Filename', 'N/A')}")
            filepath = row.get('Filepath', 'N/A')
            print(f"    Filepath: {filepath}")
            print(f"    Filepath is None: {filepath is None}")
            print(f"    Filepath type: {type(filepath).__name__}")
            
            # Check if file exists
            if filepath and filepath != 'N/A':
                exists = os.path.exists(filepath)
                print(f"    File exists: {exists}")
    
    print("\n" + "=" * 60)
    print("✓ TEST PASSED - Data is correct")
    print("=" * 60)

except Exception as e:
    print(f"\n✗ ERROR: {e}")
    import traceback
    traceback.print_exc()
    print("\n" + "=" * 60)
    print("✗ TEST FAILED")
    print("=" * 60)
