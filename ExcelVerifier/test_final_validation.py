#!/usr/bin/env python3
"""Final validation test - check all components of approved reports functionality"""

import sys
import os

# Path fix
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

import pandas as pd
from ExcelVerifier.core.database_handler import DatabaseHandler
from ExcelVerifier.config import DATABASE_FILE, APPROVED_DIRECTORY

print("=" * 80)
print("FINAL VALIDATION TEST - APPROVED REPORTS FUNCTIONALITY")
print("=" * 80)

test_results = []

# Test 1: Database connectivity
print("\n[Test 1] Database connectivity...")
try:
    db = DatabaseHandler(DATABASE_FILE)
    records = db.get_all_approved_records()
    assert len(records) > 0, "No approved records in database"
    test_results.append(("Database connectivity", True, ""))
    print(f"  ✓ PASS - {len(records)} approved records found")
except Exception as e:
    test_results.append(("Database connectivity", False, str(e)))
    print(f"  ✗ FAIL - {e}")

# Test 2: APPROVED_DIRECTORY config
print("\n[Test 2] APPROVED_DIRECTORY configuration...")
try:
    assert APPROVED_DIRECTORY, "APPROVED_DIRECTORY not configured"
    assert os.path.exists(APPROVED_DIRECTORY), f"APPROVED_DIRECTORY doesn't exist: {APPROVED_DIRECTORY}"
    test_results.append(("APPROVED_DIRECTORY config", True, ""))
    print(f"  ✓ PASS - Directory exists: {APPROVED_DIRECTORY}")
except Exception as e:
    test_results.append(("APPROVED_DIRECTORY config", False, str(e)))
    print(f"  ✗ FAIL - {e}")

# Test 3: File existence check
print("\n[Test 3] Approved files existence...")
try:
    records = db.get_all_approved_records()
    all_exist = True
    for record in records:
        filepath = record.get('filepath', '')
        if not os.path.exists(filepath):
            print(f"  ✗ File missing: {filepath}")
            all_exist = False
    
    assert all_exist, "Some approved files are missing"
    test_results.append(("File existence check", True, ""))
    print(f"  ✓ PASS - All approved files exist")
except Exception as e:
    test_results.append(("File existence check", False, str(e)))
    print(f"  ✗ FAIL - {e}")

# Test 4: Database schema
print("\n[Test 4] Database schema...")
try:
    records = db.get_all_approved_records()
    required_fields = ['id', 'order_id', 'date', 'filename', 'filepath', 
                       'company_name', 'company_nip']
    for record in records:
        for field in required_fields:
            assert field in record, f"Missing field: {field}"
    
    test_results.append(("Database schema", True, ""))
    print(f"  ✓ PASS - All required fields present")
except Exception as e:
    test_results.append(("Database schema", False, str(e)))
    print(f"  ✗ FAIL - {e}")

# Test 5: DataFrame processing
print("\n[Test 5] DataFrame processing (dialog simulation)...")
try:
    records = db.get_all_approved_records()
    df = pd.DataFrame(records)
    
    # Test rename
    df_renamed = df.rename(columns={
        'date': 'Date',
        'company_name': 'Company',
        'company_nip': 'NIP',
        'filename': 'Filename',
        'filepath': 'Filepath'
    }, errors='ignore')
    
    # Check that renamed columns exist
    expected_cols = ['Date', 'Filename', 'Filepath', 'Company', 'NIP']
    for col in expected_cols:
        assert col in df_renamed.columns, f"Missing column after rename: {col}"
    
    test_results.append(("DataFrame processing", True, ""))
    print(f"  ✓ PASS - DataFrame rename successful")
except Exception as e:
    test_results.append(("DataFrame processing", False, str(e)))
    print(f"  ✗ FAIL - {e}")

# Test 6: Filepath retrieval
print("\n[Test 6] Filepath retrieval from DataFrame...")
try:
    df = pd.DataFrame(records)
    df = df.rename(columns={'filepath': 'Filepath'}, errors='ignore')
    
    for idx, row in df.iterrows():
        filepath = row.get('Filepath', '')
        assert filepath is not None, f"Row {idx}: filepath is None"
        assert isinstance(filepath, str), f"Row {idx}: filepath is not string (type: {type(filepath)})"
        assert len(filepath) > 0, f"Row {idx}: filepath is empty"
    
    test_results.append(("Filepath retrieval", True, ""))
    print(f"  ✓ PASS - All filepaths retrieved correctly")
except Exception as e:
    test_results.append(("Filepath retrieval", False, str(e)))
    print(f"  ✗ FAIL - {e}")

# Print summary
print("\n" + "=" * 80)
print("TEST SUMMARY")
print("=" * 80)

passed = sum(1 for _, result, _ in test_results if result)
total = len(test_results)

for test_name, result, error in test_results:
    status = "✓ PASS" if result else "✗ FAIL"
    print(f"{status}: {test_name}")
    if error:
        print(f"       Error: {error}")

print(f"\nTotal: {passed}/{total} tests passed")
print("=" * 80)

if passed == total:
    print("✓ ALL TESTS PASSED - Approved reports functionality is ready!")
    sys.exit(0)
else:
    print("✗ SOME TESTS FAILED - Review errors above")
    sys.exit(1)
