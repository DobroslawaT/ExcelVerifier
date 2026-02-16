#!/usr/bin/env python3
# Patch to fix the off-by-one day error (include both start and end dates)

file_path = r"c:\Users\dobro\Downloads\Tomsystem prototyp\ExcelVerifier\ExcelVerifier\core\excel_handler.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: days_in_range calculation (add 1 to include both start and end)
content = content.replace(
    "                # Create row from current_start to report_date\n                days_in_range = (report_date - current_start).days",
    "                # Create row from current_start to report_date (inclusive of both dates)\n                days_in_range = (report_date - current_start).days + 1"
)

# Fix 2: days_to_end calculation (add 1 to include both start and end)
content = content.replace(
    "            # Final row: from last report to end of month\n            last_row = group.iloc[-1]\n            days_to_end = (last_day_of_month - current_start).days",
    "            # Final row: from last report to end of month (inclusive of both dates)\n            last_row = group.iloc[-1]\n            days_to_end = (last_day_of_month - current_start).days + 1"
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Fixed off-by-one day error")
print("Changes:")
print("  - days_in_range now includes both start and end dates (+1)")
print("  - days_to_end now includes both start and end dates (+1)")
print("  - Month start day is now properly counted")
