#!/usr/bin/env python3
# Patch to fix overlapping dates - move to next day after report date

file_path = r"c:\Users\dobro\Downloads\Tomsystem prototyp\ExcelVerifier\ExcelVerifier\core\excel_handler.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix: After creating a row ending at report_date, the next row should start the day after
old_line = "                current_start = report_date"
new_line = "                current_start = report_date + pd.Timedelta(days=1)"

content = content.replace(old_line, new_line)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Fixed overlapping date ranges")
print("Changes:")
print("  - Next period now starts day after report date")
print("  - Example: Row 1 ends Jan 16 → Row 2 starts Jan 17")
print("  - Jan 1-16 (16 days), Jan 17-31 (15 days)")
