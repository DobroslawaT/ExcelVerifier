#!/usr/bin/env python3
"""Test to check if there's an issue with how filepaths are handled when retrieved from QTableWidgetItem"""

import sys
import os

# Path fix
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

try:
    from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem
    from PyQt5.QtCore import Qt
    
    print("\nTesting QTableWidgetItem filepath handling...")
    print("=" * 60)
    
    app = QApplication([])
    
    # Create a simple table
    table = QTableWidget(1, 4)
    
    # Test filepath with backslashes
    test_filepath = r"C:\Users\dobro\Documents\Projekt zespołowy\Reports\Zatwierdzone\2026-01-03_ABCDE SP_ZO_O_ 12-123KORCZOWA.xlsx"
    print(f"\nTest filepath: {test_filepath}")
    print(f"Filepath type: {type(test_filepath).__name__}")
    print(f"Filepath length: {len(test_filepath)}")
    
    # Create QTableWidgetItem with the filepath
    item = QTableWidgetItem(test_filepath)
    print(f"\nQTableWidgetItem created")
    print(f"Item type: {type(item).__name__}")
    print(f"Item text: {item.text()}")
    print(f"Item text type: {type(item.text()).__name__}")
    
    # Set it in the table
    table.setItem(0, 3, item)
    print(f"\nItem set in table at (0, 3)")
    
    # Retrieve it from the table
    retrieved_item = table.item(0, 3)
    print(f"\nRetrieved item: {retrieved_item}")
    print(f"Retrieved item type: {type(retrieved_item).__name__}")
    
    if retrieved_item:
        retrieved_text = retrieved_item.text()
        print(f"Retrieved text: {retrieved_text}")
        print(f"Retrieved text type: {type(retrieved_text).__name__}")
        print(f"Retrieved text == original: {retrieved_text == test_filepath}")
        print(f"os.path.exists check:")
        exists = os.path.exists(retrieved_text)
        print(f"  Result: {exists}")
    
    print("\n" + "=" * 60)
    print("✓ QTableWidgetItem test passed")
    
except Exception as e:
    print(f"\n✗ Error: {e}")
    import traceback
    traceback.print_exc()
