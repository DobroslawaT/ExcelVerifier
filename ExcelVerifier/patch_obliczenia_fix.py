#!/usr/bin/env python3
# Patch script to fix Obliczenia sheet logic with month-start/end boundaries

file_path = r"c:\Users\dobro\Downloads\Tomsystem prototyp\ExcelVerifier\ExcelVerifier\core\excel_handler.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# The new logic for Sheet 2
new_sheet2_logic = '''
        # Sheet 2: Obliczenia (Calculated data with date ranges covering entire month)
        # For each company/product/month, create rows from month-start to month-end
        from calendar import monthrange
        
        calc_rows = []
        
        for (odbiorca, nazwa, miesiac), group in df.groupby(['Odbiorca', 'Nazwa', 'Miesiąc']):
            # Parse month to get first and last day
            year, month = map(int, miesiac.split('-'))
            first_day_of_month = pd.Timestamp(year=year, month=month, day=1)
            last_day_of_month = pd.Timestamp(year=year, month=month, day=monthrange(year, month)[1])
            
            # Sort by date
            group = group.sort_values('Data wystawienia').reset_index(drop=True)
            
            # Create intermediate rows
            current_start = first_day_of_month
            
            for idx, row in group.iterrows():
                report_date = row['Data wystawienia']
                
                # Create row from current_start to report_date
                days_in_range = (report_date - current_start).days
                if idx == 0:
                    # First row of month: use stan poprzedni
                    stan = row['stan poprzedni']
                else:
                    # Subsequent rows: use stan po wymianie from previous record
                    prev_row = group.loc[idx - 1]
                    stan = prev_row['stan po wymianie']
                
                butlo_dni = stan * days_in_range if days_in_range > 0 else 0
                
                calc_rows.append({
                    'Odbiorca': odbiorca,
                    'Nazwa': nazwa,
                    'Nr dokumentu': row['Nr dokumentu'] if idx == 0 else group.loc[idx - 1, 'Nr dokumentu'],
                    'Data początkowa': current_start,
                    'Data końcowa': report_date,
                    'liczba dni': days_in_range,
                    'Stan': stan,
                    'Butlo-dni': butlo_dni,
                    'Miesiąc': miesiac
                })
                
                current_start = report_date
            
            # Final row: from last report to end of month
            last_row = group.iloc[-1]
            days_to_end = (last_day_of_month - current_start).days
            final_stan = last_row['stan po wymianie']
            final_butlo = final_stan * days_to_end if days_to_end > 0 else 0
            
            calc_rows.append({
                'Odbiorca': odbiorca,
                'Nazwa': nazwa,
                'Nr dokumentu': last_row['Nr dokumentu'],
                'Data początkowa': current_start,
                'Data końcowa': last_day_of_month,
                'liczba dni': days_to_end,
                'Stan': final_stan,
                'Butlo-dni': final_butlo,
                'Miesiąc': miesiac
            })
        
        df_calc = pd.DataFrame(calc_rows)
        
        # Reorder columns
        calc_cols = [
            'Odbiorca', 'Nazwa', 'Nr dokumentu', 'Data początkowa', 'Data końcowa',
            'liczba dni', 'Stan', 'Butlo-dni', 'Miesiąc'
        ]
        df_calc = df_calc[calc_cols]
'''

# Find and replace the old Sheet 2 logic
old_section = """        # Sheet 2: Obliczenia (Calculated data with date ranges)
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
        df_calc = df_calc[calc_cols]"""

if old_section in content:
    content = content.replace(old_section, new_sheet2_logic)
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("✓ Obliczenia sheet logic updated successfully")
    print("Changes:")
    print("  - Creates rows covering entire month (start to end)")
    print("  - Month-start to first report: uses stan poprzedni")
    print("  - Between consecutive reports: uses previous stan po wymianie")
    print("  - Last report to month-end: uses last stan po wymianie")
else:
    print("✗ Could not find old section to replace")
    print("Checking for partial match...")
    if "Sheet 2: Obliczenia" in content:
        print("Found Sheet 2 comment")
    else:
        print("Sheet 2 comment not found")
