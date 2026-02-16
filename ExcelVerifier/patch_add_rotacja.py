#!/usr/bin/env python3
# Patch to add 'rotacja' column (sum of Ilość zwrócona) to Podsumowanie sheet

file_path = r"c:\Users\dobro\Downloads\Tomsystem prototyp\ExcelVerifier\ExcelVerifier\core\excel_handler.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace the Podsumowanie section to include rotacja
old_summary = """            # Arkusz 3: Podsumowanie (Summary by company/product/month)
            summary = df_calc.groupby(['Odbiorca', 'Nazwa', 'Miesiąc'])[['Butlo-dni']].sum().reset_index()
            summary.to_excel(writer, sheet_name='Podsumowanie', index=False)"""

new_summary = """            # Arkusz 3: Podsumowanie (Summary by company/product/month)
            # Calculate Butlo-dni (from calc sheet)
            summary_butlo = df_calc.groupby(['Odbiorca', 'Nazwa', 'Miesiąc'])[['Butlo-dni']].sum().reset_index()
            # Calculate rotacja (from raw sheet - sum of Ilość zwrócona)
            summary_rotacja = df_raw.groupby(['Odbiorca', 'Nazwa', 'Miesiąc'])[['Ilość zwrócona']].sum().reset_index()
            # Merge both summaries
            summary = summary_butlo.merge(summary_rotacja, on=['Odbiorca', 'Nazwa', 'Miesiąc'], how='left')
            # Rename for clarity
            summary = summary.rename(columns={'Ilość zwrócona': 'rotacja'})
            summary.to_excel(writer, sheet_name='Podsumowanie', index=False)"""

content = content.replace(old_summary, new_summary)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Added 'rotacja' column to Podsumowanie sheet")
print("Changes:")
print("  - Podsumowanie now includes: Odbiorca, Nazwa, Miesiąc, Butlo-dni, rotacja")
print("  - rotacja = sum of 'Ilość zwrócona' per company/product/month")
