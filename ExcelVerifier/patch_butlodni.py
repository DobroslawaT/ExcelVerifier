#!/usr/bin/env python3
# Patch script to update Butlo-dni calculation logic

file_path = r"c:\Users\dobro\Downloads\Tomsystem prototyp\ExcelVerifier\ExcelVerifier\core\excel_handler.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Find and replace the old logic
old_block = """        # E. Obliczanie Butlo-dni
        # Stan magazynowy * ile dni tam leżał
        df['Butlo-dni'] = df['stan po wymianie'] * df['Liczba dni']

        # Dodatkowa kolumna Miesiąc-Rok do łatwego filtrowania w Excelu
        df['Miesiąc'] = df['Data wystawienia'].dt.strftime('%Y-%m')"""

new_block = """        # E. Obliczanie Butlo-dni
        # Dodatkowa kolumna Miesiąc-Rok do łatwego filtrowania w Excelu
        df['Miesiąc'] = df['Data wystawienia'].dt.strftime('%Y-%m')

        # Identyfikacja pierwszego rekordu każdego miesiąca dla każdej pary (Odbiorca, Nazwa)
        df['Is_first_of_month'] = df.groupby(['Odbiorca', 'Nazwa', 'Miesiąc']).cumcount() == 0

        # Obliczanie Butlo-dni z inną logiką dla pierwszego rekordu miesiąca:
        # - Dla pierwszego rekordu: stan poprzedni * Liczba dni
        # - Dla pozostałych: stan po wymianie * Liczba dni
        df['Butlo-dni'] = df.apply(
            lambda row: row['stan poprzedni'] * row['Liczba dni'] if row['Is_first_of_month'] else row['stan po wymianie'] * row['Liczba dni'],
            axis=1
        )"""

if old_block in content:
    content = content.replace(old_block, new_block)
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("✓ Patch applied successfully")
else:
    print("✗ Old block not found in file")
    print("Looking for partial match...")
    if "df['Butlo-dni'] = df['stan po wymianie'] * df['Liczba dni']" in content:
        print("Found the Butlo-dni calculation line")
    else:
        print("Could not find Butlo-dni calculation")
