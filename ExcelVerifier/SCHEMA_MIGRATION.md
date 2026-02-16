# Database Schema Migration Guide

## Overview

The database has been restructured from a flat schema to a properly normalized relational design with foreign keys.

## New Schema Structure

```
companies
├── id (PK)
├── name (UNIQUE)
├── nip
├── created_at
└── updated_at

products
├── id (PK)
├── name
└── code

orders
├── id (PK)
├── company_id (FK → companies.id)
├── date_issued
└── document_number

approved_records
├── id (PK)
├── order_id (FK → orders.id)
├── date
├── filename (UNIQUE)
├── filepath
├── created_at
└── updated_at

order_items
├── id (PK)
├── order_id (FK → orders.id)
├── product_id (FK → products.id)
├── quantity_delivery
├── quantity_return
├── previous_state
├── state_after
└── created_at
```

## Relationships

- **companies** ─┬─ one-to-many → **orders**
               └─ one-to-many → **order_items** (indirect via orders)

- **orders** ─┬─ one-to-one → **approved_records**
             └─ one-to-many → **order_items**

- **products** ─── one-to-many → **order_items**

## Migration Process

### 1. Backup Your Database

The migration script automatically creates a backup, but you can manually backup:

```powershell
Copy-Item excelverifier.db excelverifier.db.backup_manual
```

### 2. Run Migration

```powershell
python migrate_to_new_schema.py
```

The migration script will:
- ✓ Create automatic backup with timestamp
- ✓ Extract companies from old approved_records
- ✓ Extract products from old reporting_data
- ✓ Map NIPs from company_db if available
- ✓ Create new normalized tables
- ✓ Migrate all transactional data
- ✓ Establish foreign key relationships
- ✓ Clean up old tables

### 3. Verify Migration

After migration, check:
```powershell
# Start the application
python ExcelVerifier/main.py

# Test:
# - View approved reports
# - Generate a new report
# - Approve a new file
```

## What Changed

### For Users
- **No visible changes** - Application works the same way
- Improved NIP search accuracy (uses normalized company data)
- Better data integrity (prevents orphaned records)

### For Developers

#### Old API (Deprecated)
```python
# Old way - flat structure
db.add_approved_record(
    date="2026-02-14",
    company="Company Name",  # ❌ Text field
    filename="file.xlsx",
    filepath="/path/to/file.xlsx"
)
```

#### New API
```python
# New way - normalized structure
company_id = db.add_company(name="Company Name", nip="1234567890")
order_id = db.add_order(
    company_id=company_id,
    date_issued="2026-02-14",
    document_number="DOC-001"
)
record_id = db.add_approved_record(
    order_id=order_id,
    date="2026-02-14",
    filename="file.xlsx",
    filepath="/path/to/file.xlsx"
)

# Add order items
products = ["Product A", "Product B"]
product_ids = [db.add_product(name=p) for p in products]

items = [
    {
        'order_id': order_id,
        'product_id': product_ids[0],
        'quantity_delivery': 100.0,
        'quantity_return': 0.0,
        'previous_state': 50.0,
        'state_after': 150.0
    },
    # ... more items
]
db.add_order_items(items)
```

## Updated Files

### Core System
- ✅ `database_handler.py` - New schema, all CRUD methods updated
- ✅ `excel_handler.py` - Approval process updated
- ✅ `dialogs.py` - Approved reports dialog updated

### Migration
- ✅ `migrate_to_new_schema.py` - Complete data migration script

### Pending Updates
- ⚠️ `import_export.py` - **Needs updating for new schema**
  - Database merge function
  - Excel import function
  - Temporarily may not work correctly

## Known Limitations

1. **Import/Export Functions**: The database import/export features need to be updated to work with the new schema. Importing from old exports may fail.

2. **Backward Compatibility**: Old database files must be migrated using the migration script. The application cannot directly open old-schema databases.

## Rollback Procedure

If you need to rollback to the old schema:

```powershell
# 1. Stop the application
# 2. Find your backup file (e.g., excelverifier.db.backup_20260214_153000)
# 3. Restore it
Copy-Item excelverifier.db.backup_TIMESTAMP excelverifier.db -Force
```

## Benefits of New Schema

1. **Data Integrity**: Foreign keys prevent orphaned records
2. **Normalization**: No duplicate company names or product names
3. **Query Performance**: Better indexed for common queries
4. **Scalability**: Easier to add features (e.g., product categories, company contacts)
5. **NIP Accuracy**: Direct company→NIP relationship, no text parsing needed
6. **Audit Trail**: Proper timestamps on all tables

## Troubleshooting

### "Table already exists" error
- You're running migration on already-migrated database
- Check if `orders` table exists: If yes, migration already done

### "No approved records found"
- Migration didn't find old data
- Check if backup file exists and contains data

### NIP values are NULL
- This is normal if company_db didn't have NIP data
- NIPs can be added later through company management

### Foreign key constraint errors
- Ensure foreign keys are enabled: `PRAGMA foreign_keys = ON`
- Check data integrity before migration

## Support

Migration creates detailed backup files. If anything goes wrong:
1. Keep the backup file safe
2. Note the exact error message
3. Check the migration script output logs
