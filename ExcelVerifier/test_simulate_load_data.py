#!/usr/bin/env python3
"""Simulate what ApprovedReportsDialog.load_data() does"""

import sys
import os
from datetime import datetime

# Path fix
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

import pandas as pd

print("\n" + "=" * 70)
print("SIMULATING APPROVEDDIALOG.LOAD_DATA()")
print("=" * 70)

try:
    print("\n[1] Importing modules...")
    from ExcelVerifier.core.database_handler import DatabaseHandler
    from ExcelVerifier.config import DATABASE_FILE
    from ExcelVerifier.core.company_db import normalize_nip
    
    print(f"[2] Opening database: {DATABASE_FILE}")
    db = DatabaseHandler(DATABASE_FILE)
    
    print("[3] Calling get_all_approved_records()...")
    records = db.get_all_approved_records()
    print(f"[4] Got {len(records)} records")
    
    print("\n[5] Creating DataFrame...")
    df = pd.DataFrame(records)
    print(f"    Shape: {df.shape}")
    print(f"    Columns: {df.columns.tolist()}")
    
    if df.empty:
        print("\n✗ DataFrame is EMPTY!")
    else:
        print(f"\n[6] DataFrame has data, renaming columns...")
        df = df.rename(columns={
            'date': 'Date',
            'company_name': 'Company',
            'company_nip': 'NIP',
            'filename': 'Filename',
            'filepath': 'Filepath'
        }, errors='ignore')
        
        print(f"    After rename: {df.columns.tolist()}")
        print(f"    'Filepath' in columns: {'Filepath' in df.columns}")
        
        if 'NIP' in df.columns:
            df['NIP'] = df['NIP'].apply(lambda x: normalize_nip(x) if x else "")
        
        # Sort by date
        df['_sort_date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.sort_values(by='_sort_date', ascending=False)
        df.drop(columns=['_sort_date'], inplace=True)
        
        # Month filtering (like the dialog does)
        filter_month = None
        current_month = datetime.now().strftime('%Y-%m')
        print(f"\n[7] Month filtering:")
        print(f"    filter_month (from parent): {filter_month}")
        print(f"    current_month: {current_month}")
        
        if filter_month:
            df_filtered = df[df['Date'].astype(str).str.startswith(filter_month)]
        else:
            df_filtered = df[df['Date'].astype(str).str.startswith(current_month)]
            
            if df_filtered.empty:
                latest_date = df['Date'].astype(str).max()
                if latest_date:
                    latest_month = latest_date[:7]
                    print(f"    No records for current month, using latest month: {latest_month}")
                    df_filtered = df[df['Date'].astype(str).str.startswith(latest_month)]
        
        if df_filtered.empty:
            print(f"    df_filtered is empty, using all records")
            df_filtered = df.copy()
        
        print(f"\n[8] Final result:")
        print(f"    Total table df: {len(df)} rows")
        print(f"    Month-filtered df: {len(df_filtered)} rows")
        
        print(f"\n[9] Displaying first row of table df:")
        if len(df) > 0:
            print(f"    {df.iloc[0].to_dict()}")
        
        print(f"\n✓ SUCCESS - Would display {len(df)} records in table")
        
except Exception as e:
    print(f"\n✗ ERROR: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 70)
