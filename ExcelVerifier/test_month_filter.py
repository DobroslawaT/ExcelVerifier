#!/usr/bin/env python
"""Test month filtering in approved reports dialog"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ExcelVerifier.core.database_handler import DatabaseHandler
from datetime import datetime

# Initialize database handler
db = DatabaseHandler()

# Get all approved records
all_records = db.get_all_approved_records()
print(f"Total approved records: {len(all_records)}")

# Show sample records with dates
print("\nSample records by date:")
dates_by_month = {}
for record in all_records:
    # Handle both tuple and dict formats
    if isinstance(record, (tuple, list)):
        date_str = record[1]  # date column
    else:
        date_str = record.get('date', record.get('Date', ''))
    
    if date_str:
        month = str(date_str)[:7]  # YYYY-MM
        if month not in dates_by_month:
            dates_by_month[month] = []
        dates_by_month[month].append(date_str)

for month in sorted(dates_by_month.keys()):
    print(f"  {month}: {len(dates_by_month[month])} records")
    print(f"    Examples: {dates_by_month[month][:3]}")

# Test filtering logic like the dialog would do it
print("\n--- Testing Dialog Filtering Logic ---")

# Simulate filtering to February 2026
test_filter_month = "2026-02"
print(f"\nFiltering to month: {test_filter_month}")

# Convert to DataFrame for filtering (like the dialog does)
import pandas as pd

# Create DataFrame - recognize the format of all_records
if all_records and isinstance(all_records[0], (tuple, list)):
    df = pd.DataFrame(all_records, columns=['id', 'Date', 'Company', 'Filename', 'Filepath', 'created_at', 'updated_at'])
else:
    # If it's already dict-like, pandas should handle it
    df = pd.DataFrame(all_records)
    if 'date' in df.columns and 'Date' not in df.columns:
        df = df.rename(columns={'date': 'Date', 'company': 'Company', 'filename': 'Filename', 'filepath': 'Filepath'})

print(f"Before filter: {len(df)} records")
print(f"Date column sample: {df['Date'].head()}")

filtered_df = df[df['Date'].astype(str).str.startswith(test_filter_month)]
print(f"After filter to {test_filter_month}: {len(filtered_df)} records")

if not filtered_df.empty:
    print("\nFiltered records:")
    for idx, row in filtered_df.iterrows():
        print(f"  {row['Date']} | {row['Company']} | {row['Filename']}")
else:
    print(f"No records found for {test_filter_month}")

print("\n✓ Test completed successfully")
