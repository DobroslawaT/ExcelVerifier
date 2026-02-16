#!/usr/bin/env python3
# Patch to fix day counting without changing display dates

file_path = r"c:\Users\dobro\Downloads\Tomsystem prototyp\ExcelVerifier\ExcelVerifier\core\excel_handler.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Revert the +1 day change
content = content.replace(
    "                current_start = report_date + pd.Timedelta(days=1)",
    "                current_start = report_date"
)

# Fix the day calculation: for non-first rows, don't count the start date (already counted in previous row)
old_calc = """                # Create row from current_start to report_date (inclusive of both dates)
                days_in_range = (report_date - current_start).days + 1
                if idx == 0:
                    # First row of month: use stan poprzedni
                    stan = row['stan poprzedni']
                else:
                    # Subsequent rows: use stan po wymianie from previous record
                    prev_row = group.loc[idx - 1]
                    stan = prev_row['stan po wymianie']"""

new_calc = """                # Create row from current_start to report_date
                # For first row: count all days including boundaries (+1)
                # For subsequent rows: don't double-count the start date (already in previous row end)
                if idx == 0:
                    # First row of month: use stan poprzedni, count all days
                    days_in_range = (report_date - current_start).days + 1
                    stan = row['stan poprzedni']
                else:
                    # Subsequent rows: use stan po wymianie from previous record, exclude overlap
                    days_in_range = (report_date - current_start).days
                    prev_row = group.loc[idx - 1]
                    stan = prev_row['stan po wymianie']"""

content = content.replace(old_calc, new_calc)

# Also fix the final row calculation
content = content.replace(
    "            # Final row: from last report to end of month (inclusive of both dates)\n            last_row = group.iloc[-1]\n            days_to_end = (last_day_of_month - current_start).days + 1",
    "            # Final row: from last report to end of month (exclude overlap with previous row)\n            last_row = group.iloc[-1]\n            days_to_end = (last_day_of_month - current_start).days"
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Fixed day calculation without changing display dates")
print("Changes:")
print("  - Display: 01.01-16.01 and 16.01-31.01 (as is)")
print("  - Days: 16 and 15 (no double-counting)")
print("  - First row: counts both boundaries")
print("  - Subsequent rows: exclude start date from count")
