#!/usr/bin/env python3
"""Quick script to check approved records in the database and export to Excel."""

import os
import sys
from datetime import datetime

# Add the project to path
sys.path.insert(0, os.path.dirname(__file__))

from ExcelVerifier.core.database_handler import DatabaseHandler
from ExcelVerifier.config import DATABASE_FILE

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

def export_to_excel(records):
    """Export approved records to an Excel file."""
    if not HAS_PANDAS:
        print("⚠️  pandas not installed - skipping Excel export (install: pip install pandas openpyxl)\n")
        return None
    
    try:
        # Create DataFrame
        df = pd.DataFrame(records)
        
        # Ensure proper column order
        columns = ['date', 'company', 'filename', 'filepath', 'created_at', 'updated_at']
        df = df[[col for col in columns if col in df.columns]]
        
        # Rename columns for display
        df.columns = ['Date', 'Company', 'Filename', 'Filepath', 'Created', 'Updated']
        
        # Format dates
        for col in ['Date', 'Created', 'Updated']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d %H:%M')
        
        # Create export filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_file = f"ApprovedRecords_Export_{timestamp}.xlsx"
        
        # Export to Excel
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Approved Records', index=False)
            
            # Format the Excel file
            workbook = writer.book
            worksheet = writer.sheets['Approved Records']
            
            # Adjust column widths
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 35
            worksheet.column_dimensions['C'].width = 30
            worksheet.column_dimensions['D'].width = 50
            worksheet.column_dimensions['E'].width = 20
            worksheet.column_dimensions['F'].width = 20
            
            # Bold header row
            from openpyxl.styles import Font
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
        
        print(f"📊 Excel file created: {excel_file}\n")
        return excel_file
        
    except Exception as e:
        print(f"⚠️  Error exporting to Excel: {str(e)}\n")
        return None

def main():

    print("\n" + "="*80)
    print("📋 APPROVED RECORDS - DATABASE CHECK")
    print("="*80 + "\n")
    
    try:
        db = DatabaseHandler(DATABASE_FILE)
        
        # Get all approved records
        records = db.get_all_approved_records()
        
        if not records:
            print("❌ No approved records found in database.\n")
            return
        
        print(f"✅ Found {len(records)} approved records:\n")
        print("-" * 80)
        print(f"{'Date':<12} | {'Company':<30} | {'Filename':<25} | {'Filepath':<20}")
        print("-" * 80)
        
        for i, record in enumerate(records, 1):
            date = record.get('date', 'N/A')[:10]
            company = str(record.get('company', 'N/A'))[:30]
            filename = str(record.get('filename', 'N/A'))[:25]
            filepath = str(record.get('filepath', 'N/A'))[:20]
            
            print(f"{date:<12} | {company:<30} | {filename:<25} | {filepath:<20}")
            
            if i % 10 == 0:
                print("-" * 80)
        
        print("-" * 80)
        print(f"\n📊 Summary:")
        print(f"  • Total records: {len(records)}")
        
        # Get unique companies
        companies = set(r.get('company', 'Unknown') for r in records)
        print(f"  • Unique companies: {len(companies)}")
        print(f"  • Companies: {', '.join(sorted(companies))}\n")
        
        # Get available months
        months = db.get_available_months()
        if months:
            print(f"  • Available months: {', '.join(months)}\n")
        
        # Get database stats
        stats = db.get_database_stats()
        print(f"  • Database Stats: {stats}\n")
        
        # Export to Excel
        export_to_excel(records)
        
        db.close()
        
    except Exception as e:
        print(f"❌ Error: {str(e)}\n")
        return 1
    
    print("="*80 + "\n")
    return 0

if __name__ == "__main__":
    sys.exit(main())
