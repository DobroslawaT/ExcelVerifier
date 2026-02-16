#!/usr/bin/env python3
"""Test script to comprehensively test the approved reports dialog workflow"""

import sys
import os

# Path fix
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

import pandas as pd
from ExcelVerifier.core.database_handler import DatabaseHandler
from ExcelVerifier.config import DATABASE_FILE

print("=" * 80)
print("COMPREHENSIVE APPROVED REPORTS DIALOG TEST")
print("=" * 80)

def test_dialog_workflow():
    """Simulate the complete dialog data loading and retrieval workflow"""
    
    # Step 1: Load data from database (as ApprovedReportsDialog.load_data does)
    print("\n[Step 1] Load approved records from database...")
    try:
        db = DatabaseHandler(DATABASE_FILE)
        records = db.get_all_approved_records()
        print(f"  ✓ Loaded {len(records)} records")
    except Exception as e:
        print(f"  ✗ FAILED: {e}")
        return False
    
    # Step 2: Create DataFrame (as load_data does)
    print("\n[Step 2] Create DataFrame from records...")
    try:
        df = pd.DataFrame(records)
        print(f"  ✓ DataFrame created: {df.shape}")
    except Exception as e:
        print(f"  ✗ FAILED: {e}")
        return False
    
    # Step 3: Rename columns (as load_data does)
    print("\n[Step 3] Rename DataFrame columns...")
    try:
        df = df.rename(columns={
            'date': 'Date',
            'company_name': 'Company',
            'company_nip': 'NIP',
            'filename': 'Filename',
            'filepath': 'Filepath'
        }, errors='ignore')
        print(f"  ✓ Columns renamed: {df.columns.tolist()}")
    except Exception as e:
        print(f"  ✗ FAILED: {e}")
        return False
    
    # Step 4: Simulate populate_table (extract filepath from each row)
    print("\n[Step 4] Simulate populate_table - extract filepaths...")
    filepaths_for_table = []
    try:
        for r, row in df.reset_index(drop=True).iterrows():
            # This simulates what populate_table does
            filepath = row.get('Filepath', '')
            
            # Ensure filepath is not None and is a valid string
            if filepath is None:
                filepath = ''
            else:
                filepath = str(filepath)
            
            print(f"  Row {r}: filepath = {filepath if len(str(filepath)) < 60 else filepath[:60] + '...'}")
            print(f"           type = {type(filepath).__name__}, is_empty = {len(filepath) == 0}")
            
            # Check if file exists (as populate_table should check)
            if filepath:
                exists = os.path.exists(filepath)
                print(f"           file_exists = {exists}")
                if not exists:
                    print(f"           WARNING: File doesn't exist!")
            
            filepaths_for_table.append(filepath)
        
        print(f"  ✓ Extracted {len(filepaths_for_table)} filepaths")
    except Exception as e:
        print(f"  ✗ FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    # Step 5: Simulate accept_selection (retrieve filepath from table)
    print("\n[Step 5] Simulate accept_selection - retrieve filepath from simulated table...")
    try:
        if filepaths_for_table:
            # Simulate user selecting first row (row 0)
            row = 0
            filepath_from_table = filepaths_for_table[row]
            
            print(f"  Retrieving filepath from row {row}...")
            print(f"      filepath = {filepath_from_table}")
            print(f"      type = {type(filepath_from_table).__name__}")
            print(f"      is_None = {filepath_from_table is None}")
            print(f"      is_empty_string = {filepath_from_table == ''}")
            
            # Check filepath validity (as accept_selection does)
            if not filepath_from_table:
                print(f"  ✗ Filepath is empty!")
                return False
            
            if not os.path.exists(filepath_from_table):
                print(f"  ✗ File doesn't exist: {filepath_from_table}")
                return False
            
            print(f"  ✓ Filepath is valid and file exists")
        else:
            print(f"  ✗ No filepaths available")
            return False
    except Exception as e:
        print(f"  ✗ FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    print("\n" + "=" * 80)
    print("✓ ALL TESTS PASSED - Workflow is correct!")
    print("=" * 80)
    return True

if __name__ == "__main__":
    success = test_dialog_workflow()
    sys.exit(0 if success else 1)
