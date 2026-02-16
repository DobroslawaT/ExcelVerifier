#!/usr/bin/env python3
# Patch to ignore negative values in rotacja calculation

file_path = r"c:\Users\dobro\Downloads\Tomsystem prototyp\ExcelVerifier\ExcelVerifier\core\excel_handler.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace the rotacja calculation to filter out negative values
old_rotacja = """            # Calculate rotacja (from raw sheet - sum of Ilość zwrócona)
            summary_rotacja = df_raw.groupby(['Odbiorca', 'Nazwa', 'Miesiąc'])[['Ilość zwrócona']].sum().reset_index()"""

new_rotacja = """            # Calculate rotacja (from raw sheet - sum of Ilość zwrócona, ignoring negatives)
            # Filter to only positive values
            df_raw_positive = df_raw[df_raw['Ilość zwrócona'] > 0].copy()
            summary_rotacja = df_raw_positive.groupby(['Odbiorca', 'Nazwa', 'Miesiąc'])[['Ilość zwrócona']].sum().reset_index()"""

content = content.replace(old_rotacja, new_rotacja)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Updated rotacja to ignore negative values")
print("Changes:")
print("  - Only positive 'Ilość zwrócona' values are summed")
print("  - Negative values are excluded from rotacja calculation")
