#!/usr/bin/env python
"""Test script with detailed logging to diagnose startup issues."""

import sys
import traceback
import os

# Write to a log file so we can see what's happening
log_file = "app_startup.log"

def log(msg):
    print(msg)
    with open(log_file, 'a') as f:
        f.write(msg + '\n')

# Clear old log
if os.path.exists(log_file):
    os.remove(log_file)

log("=" * 70)
log("ExcelVerifier Startup Diagnostics")
log("=" * 70)

try:
    log("\n[1] Importing PyQt5...")
    from PyQt5.QtWidgets import QApplication
    log("✓ PyQt5 imported")
    
    log("\n[2] Setting up path...")
    sys.path.insert(0, 'ExcelVerifier')
    log("✓ Path setup complete")
    
    log("\n[3] Importing ui.main_window...")
    from ui.main_window import VerifyApp
    log("✓ VerifyApp imported")
    
    log("\n[4] Creating QApplication...")
    app = QApplication(sys.argv)
    log("✓ QApplication created")
    
    log("\n[5] Creating VerifyApp window...")
    window = VerifyApp()
    log("✓ VerifyApp window created")
    
    log("\n[6] Showing window maximized...")
    window.showMaximized()
    log("✓ Window shown")
    
    log("\n[7] Starting event loop...")
    sys.exit(app.exec_())
    
except Exception as e:
    log(f"\n❌ ERROR: {e}")
    log("\nFull Traceback:")
    log(traceback.format_exc())
    sys.exit(1)
