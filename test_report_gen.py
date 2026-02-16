#!/usr/bin/env python
from ExcelVerifier.core.excel_handler import ExcelHandler
from ExcelVerifier.core.file_manager import FileManager
import traceback

try:
    excel_handler = ExcelHandler()
    filters = {
        'mode': 1,
        'month': '2026-01',
        'from_date': None,
        'to_date': None,
        'company': None
    }
    result = excel_handler.generate_report(filters, 'test_report.xlsx')
    print(f"\n\nReport generated: {result}")
except Exception as e:
    print(f"\n\nError: {e}")
    traceback.print_exc()
