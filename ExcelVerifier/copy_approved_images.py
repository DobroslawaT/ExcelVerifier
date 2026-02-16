#!/usr/bin/env python3
"""
Script to copy images for already-approved reports to the Zatwierdzone folder
"""

import sys
import os
import shutil

# Path fix
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

from ExcelVerifier.core.database_handler import DatabaseHandler
from ExcelVerifier.config import DATABASE_FILE, REPORTS_ROOT, APPROVED_DIRECTORY

print("=" * 80)
print("COPY IMAGES FOR APPROVED REPORTS")
print("=" * 80)

try:
    # Load all approved records
    db = DatabaseHandler(DATABASE_FILE)
    records = db.get_all_approved_records()
    print(f"\nFound {len(records)} approved records in database")
    
    copied_count = 0
    not_found_count = 0
    
    for record in records:
        filepath = record['filepath']
        filename = record['filename']
        
        # Get base name without extension
        excel_base = os.path.splitext(filename)[0]
        
        # Look for image in Niezatwierdzone folder
        niezatwierdzone_dir = os.path.join(os.path.dirname(APPROVED_DIRECTORY), 'Niezatwierdzone')
        
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']
        found_image = False
        
        for ext in image_extensions:
            source_image = os.path.join(niezatwierdzone_dir, excel_base + ext)
            if os.path.exists(source_image):
                dest_image = os.path.join(APPROVED_DIRECTORY, excel_base + ext)
                
                # Check if destination already exists
                if os.path.exists(dest_image):
                    print(f"  ⏭️  {excel_base}{ext} - already exists in Zatwierdzone")
                else:
                    # Copy the image
                    try:
                        shutil.copy2(source_image, dest_image)
                        print(f"  ✓ {excel_base}{ext} - copied to Zatwierdzone")
                        copied_count += 1
                    except Exception as e:
                        print(f"  ✗ {excel_base}{ext} - error copying: {e}")
                
                found_image = True
                break
        
        if not found_image:
            print(f"  ⚠️  {excel_base} - no image found in Niezatwierdzone")
            not_found_count += 1
    
    print(f"\n" + "=" * 80)
    print(f"Summary:")
    print(f"  Images copied: {copied_count}")
    print(f"  Images not found: {not_found_count}")
    print(f"=" * 80)
    
except Exception as e:
    print(f"\n✗ ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
