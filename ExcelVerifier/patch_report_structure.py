#!/usr/bin/env python3
# Patch script to restructure Excel report with two sheets

file_path = r"c:\Users\dobro\Downloads\Tomsystem prototyp\ExcelVerifier\ExcelVerifier\core\excel_handler.py"

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the line number where the export section starts
export_start = None
for i, line in enumerate(lines):
    if "# --- 4. EKSPORT DO EXCELA ---" in line:
        export_start = i
        break

if export_start is None:
    print("Could not find export section")
    exit(1)

# Replace the export section
old_section_end = None
for i in range(export_start, len(lines)):
    if "return output_path" in lines[i]:
        old_section_end = i + 1
        break

new_export_code = '''        # --- 4. EKSPORT DO EXCELA ---

        # Sheet 1: Dane Szczegółowe (Raw data - no calculations)
        raw_cols = [
            'Odbiorca', 'Nazwa', 'Nr dokumentu', 'Data wystawienia',
            'Ilość zamówiona', 'Ilość zwrócona',
            'stan poprzedni', 'stan po wymianie', 'Miesiąc'
        ]
        df_raw = df[raw_cols].copy()

        # Sheet 2: Obliczenia (Calculated data with date ranges)
        # Group by (Odbiorca, Nazwa) to create date ranges
        df_calc = df[['Odbiorca', 'Nazwa', 'Nr dokumentu', 'Data wystawienia', 
                      'Data następna', 'Liczba dni', 'stan po wymianie', 'Butlo-dni', 'Miesiąc', 'Is_first_of_month']].copy()
        
        # Rename for clarity
        df_calc = df_calc.rename(columns={
            'Data wystawienia': 'Data początkowa',
            'Data następna': 'Data końcowa',
            'stan po wymianie': 'Stan',
            'Liczba dni': 'liczba dni'
        })

        # Reorder columns
        calc_cols = [
            'Odbiorca', 'Nazwa', 'Nr dokumentu', 'Data początkowa', 'Data końcowa',
            'liczba dni', 'Stan', 'Butlo-dni', 'Miesiąc', 'Is_first_of_month'
        ]
        df_calc = df_calc[calc_cols]

        # Ścieżka zapisu
        output_filename = f"Raport_ButloDni_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        output_path = os.path.join(os.path.dirname(APPROVED_FILE), output_filename)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Arkusz 1: Dane Szczegółowe (Raw data)
            df_raw.to_excel(writer, sheet_name='Dane Szczegółowe', index=False)
            
            # Formatowanie dat w Arkuszu 1
            ws1 = writer.sheets['Dane Szczegółowe']
            for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row):
                # Kolumna D (4) to Data wystawienia
                row[3].number_format = 'dd.mm.yyyy'

            # Arkusz 2: Obliczenia (Calculated data with date ranges)
            df_calc.to_excel(writer, sheet_name='Obliczenia', index=False)
            
            # Formatowanie dat w Arkuszu 2
            ws2 = writer.sheets['Obliczenia']
            for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
                # Kolumna D (4) to Data początkowa
                row[3].number_format = 'dd.mm.yyyy'
                # Kolumna E (5) to Data końcowa
                row[4].number_format = 'dd.mm.yyyy'

            # Arkusz 3: Podsumowanie (Summary by company/product/month)
            summary = df_calc.groupby(['Odbiorca', 'Nazwa', 'Miesiąc'])[['Butlo-dni']].sum().reset_index()
            summary.to_excel(writer, sheet_name='Podsumowanie', index=False)

            # Formatowanie szerokości kolumn (kosmetyka)
            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                for column in ws.columns:
                    max_length = 0
                    column = list(column)
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

        return output_path
'''

# Replace old section with new
lines = lines[:export_start] + [new_export_code + '\n'] + lines[old_section_end:]

with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("✓ Report structure updated successfully")
print("Changes:")
print("  - Sheet 1 (Dane Szczegółowe): Raw data only")
print("  - Sheet 2 (Obliczenia): Date ranges and calculations")
print("  - Sheet 3 (Podsumowanie): Summary by company/product/month")
