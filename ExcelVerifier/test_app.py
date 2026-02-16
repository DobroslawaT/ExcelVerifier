#!/usr/bin/env python
"""Test script to debug app startup issues."""

import sys
import traceback

try:
    print("Importing PyQt5...")
    from PyQt5.QtWidgets import QApplication
    print("✓ PyQt5 imported")
    
    print("Importing ui.main_window...")
    sys.path.insert(0, 'ExcelVerifier')
    from ui.main_window import VerifyApp
    print("✓ VerifyApp imported")
    
    print("Creating QApplication...")
    app = QApplication(sys.argv)
    print("✓ QApplication created")
    
    print("Creating VerifyApp window...")
    window = VerifyApp()
    print("✓ VerifyApp window created")
    
    print("Showing window...")
    window.showMaximized()
    print("✓ Window shown")
    
    print("Starting app loop...")
    sys.exit(app.exec_())
    
except Exception as e:
    print(f"\n❌ ERROR: {e}")
    traceback.print_exc()
    sys.exit(1)
