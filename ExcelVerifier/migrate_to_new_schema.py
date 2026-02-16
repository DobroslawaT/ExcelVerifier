"""
Migration script to convert database from old schema to new normalized structure.

Old schema:
- approved_records (date, company, filename, filepath)
- reporting_data (date_issued, recipient, product_name, quantities, source_filename)

New schema:
- companies (id, name, nip)
- products (id, name, code)
- orders (id, company_id, date_issued, document_number)
- approved_records (id, order_id, date, filename, filepath)
- order_items (id, order_id, product_id, quantities)
"""

import sqlite3
import shutil
import os
from datetime import datetime
from typing import Dict, List, Set


def backup_database(db_path: str) -> str:
    """Create a backup of the database before migration."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{db_path}.backup_{timestamp}"
    shutil.copy2(db_path, backup_path)
    print(f"✓ Database backed up to: {backup_path}")
    return backup_path


def check_old_schema_exists(conn: sqlite3.Connection) -> bool:
    """Check if old reporting_data table exists."""
    cursor = conn.cursor()
    cursor.execute("""
        SELECT name FROM sqlite_master 
        WHERE type='table' AND name='reporting_data'
    """)
    return cursor.fetchone() is not None


def get_old_schema_data(conn: sqlite3.Connection) -> Dict:
    """Retrieve all data from old schema."""
    cursor = conn.cursor()
    
    # Get old approved_records structure
    cursor.execute("""
        SELECT name FROM pragma_table_info('approved_records')
    """)
    columns = [row[0] for row in cursor.fetchall()]
    has_old_structure = 'company' in columns
    
    if not has_old_structure:
        print("✓ Database already has new schema structure")
        return None
    
    cursor.execute("SELECT * FROM approved_records")
    old_approved = [dict(zip([d[0] for d in cursor.description], row)) 
                    for row in cursor.fetchall()]
    
    cursor.execute("SELECT * FROM reporting_data")
    old_reporting = [dict(zip([d[0] for d in cursor.description], row)) 
                     for row in cursor.fetchall()]
    
    print(f"✓ Found {len(old_approved)} approved records")
    print(f"✓ Found {len(old_reporting)} reporting data records")
    
    return {
        'approved_records': old_approved,
        'reporting_data': old_reporting
    }


def extract_companies(old_data: Dict) -> List[Dict]:
    """Extract unique companies from old data."""
    companies = {}
    
    # From approved_records
    for record in old_data['approved_records']:
        company_name = record.get('company', '').strip()
        if company_name and company_name not in companies:
            companies[company_name] = {
                'name': company_name,
                'nip': None  # NIP will be populated from company_db or user input
            }
    
    # From reporting_data
    for record in old_data['reporting_data']:
        recipient = record.get('recipient', '').strip()
        if recipient and recipient not in companies:
            companies[recipient] = {
                'name': recipient,
                'nip': None
            }
    
    company_list = list(companies.values())
    print(f"✓ Extracted {len(company_list)} unique companies")
    return company_list


def extract_products(old_data: Dict) -> List[Dict]:
    """Extract unique products from reporting data."""
    products = set()
    
    for record in old_data['reporting_data']:
        product_name = record.get('product_name', '').strip()
        if product_name:
            products.add(product_name)
    
    product_list = [{'name': name, 'code': None} for name in sorted(products)]
    print(f"✓ Extracted {len(product_list)} unique products")
    return product_list


def create_new_schema(conn: sqlite3.Connection):
    """Drop old tables and create new schema."""
    cursor = conn.cursor()
    
    print("→ Dropping old tables...")
    cursor.execute("DROP TABLE IF EXISTS reporting_data")
    
    # Rename old approved_records to backup
    cursor.execute("ALTER TABLE approved_records RENAME TO approved_records_old")
    
    print("→ Creating new schema tables...")
    
    # Companies table
    cursor.execute("""
        CREATE TABLE companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            nip TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cursor.execute("CREATE INDEX idx_companies_name ON companies(name)")
    cursor.execute("CREATE INDEX idx_companies_nip ON companies(nip)")
    
    # Products table
    cursor.execute("""
        CREATE TABLE products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            code TEXT,
            UNIQUE(name, code)
        )
    """)
    cursor.execute("CREATE INDEX idx_products_name ON products(name)")
    
    # Orders table
    cursor.execute("""
        CREATE TABLE orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            date_issued TEXT NOT NULL,
            document_number TEXT,
            FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE
        )
    """)
    cursor.execute("CREATE INDEX idx_orders_company ON orders(company_id)")
    cursor.execute("CREATE INDEX idx_orders_date ON orders(date_issued)")
    
    # New approved_records table
    cursor.execute("""
        CREATE TABLE approved_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            filename TEXT UNIQUE NOT NULL,
            filepath TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE
        )
    """)
    cursor.execute("CREATE INDEX idx_approved_date ON approved_records(date)")
    cursor.execute("CREATE INDEX idx_approved_filename ON approved_records(filename)")
    cursor.execute("CREATE INDEX idx_approved_order ON approved_records(order_id)")
    
    # Order items table
    cursor.execute("""
        CREATE TABLE order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            quantity_delivery REAL DEFAULT 0,
            quantity_return REAL DEFAULT 0,
            previous_state REAL DEFAULT 0,
            state_after REAL DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE,
            FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE
        )
    """)
    cursor.execute("CREATE INDEX idx_order_items_order ON order_items(order_id)")
    cursor.execute("CREATE INDEX idx_order_items_product ON order_items(product_id)")
    
    conn.commit()
    print("✓ New schema created successfully")


def populate_companies(conn: sqlite3.Connection, companies: List[Dict]) -> Dict[str, int]:
    """Insert companies and return name->id mapping."""
    cursor = conn.cursor()
    company_map = {}
    
    for company in companies:
        cursor.execute("""
            INSERT INTO companies (name, nip)
            VALUES (?, ?)
        """, (company['name'], company['nip']))
        company_map[company['name']] = cursor.lastrowid
    
    conn.commit()
    print(f"✓ Inserted {len(company_map)} companies")
    return company_map


def populate_products(conn: sqlite3.Connection, products: List[Dict]) -> Dict[str, int]:
    """Insert products and return name->id mapping."""
    cursor = conn.cursor()
    product_map = {}
    
    for product in products:
        cursor.execute("""
            INSERT INTO products (name, code)
            VALUES (?, ?)
        """, (product['name'], product['code']))
        product_map[product['name']] = cursor.lastrowid
    
    conn.commit()
    print(f"✓ Inserted {len(product_map)} products")
    return product_map


def migrate_data(conn: sqlite3.Connection, old_data: Dict, 
                 company_map: Dict[str, int], product_map: Dict[str, int]):
    """Migrate old data to new schema structure."""
    cursor = conn.cursor()
    
    # Group reporting_data by source_filename to create orders
    orders_by_file = {}
    for record in old_data['reporting_data']:
        filename = record.get('source_filename', '')
        if filename not in orders_by_file:
            orders_by_file[filename] = []
        orders_by_file[filename].append(record)
    
    print(f"→ Creating {len(orders_by_file)} orders from reporting data...")
    
    order_map = {}  # filename -> order_id mapping
    
    for filename, items in orders_by_file.items():
        # Find corresponding approved_record for company info
        approved = None
        for rec in old_data['approved_records']:
            if rec['filename'] == filename:
                approved = rec
                break
        
        if not approved:
            print(f"  ⚠ Warning: No approved record found for {filename}, skipping...")
            continue
        
        # Get company_id
        company_name = approved.get('company', '')
        company_id = company_map.get(company_name)
        
        if not company_id:
            print(f"  ⚠ Warning: Company '{company_name}' not found, skipping {filename}")
            continue
        
        # Use first item's date_issued for order date
        date_issued = items[0].get('date_issued', approved['date'])
        document_number = items[0].get('document_number')
        
        # Create order
        cursor.execute("""
            INSERT INTO orders (company_id, date_issued, document_number)
            VALUES (?, ?, ?)
        """, (company_id, date_issued, document_number))
        order_id = cursor.lastrowid
        order_map[filename] = order_id
        
        # Create approved_record linked to this order
        cursor.execute("""
            INSERT INTO approved_records (order_id, date, filename, filepath, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (order_id, approved['date'], approved['filename'], approved['filepath'],
              approved.get('created_at'), approved.get('updated_at')))
        
        # Create order_items
        for item in items:
            product_name = item.get('product_name', '').strip()
            if not product_name:
                continue
            
            product_id = product_map.get(product_name)
            if not product_id:
                print(f"  ⚠ Warning: Product '{product_name}' not found")
                continue
            
            cursor.execute("""
                INSERT INTO order_items 
                (order_id, product_id, quantity_delivery, quantity_return, 
                 previous_state, state_after, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (order_id, product_id,
                  item.get('quantity_delivery', 0),
                  item.get('quantity_return', 0),
                  item.get('previous_state', 0),
                  item.get('state_after', 0),
                  item.get('created_at')))
    
    conn.commit()
    print(f"✓ Migrated {len(order_map)} orders with their items and approved records")


def cleanup_old_tables(conn: sqlite3.Connection):
    """Remove old backup tables after successful migration."""
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS approved_records_old")
    conn.commit()
    print("✓ Cleaned up old tables")


def main():
    """Main migration function."""
    db_path = "excelverifier.db"
    
    if not os.path.exists(db_path):
        print(f"✗ Database not found: {db_path}")
        return
    
    print("\n" + "="*60)
    print("DATABASE MIGRATION: Old Schema → New Normalized Schema")
    print("="*60 + "\n")
    
    # Create backup
    backup_path = backup_database(db_path)
    
    try:
        conn = sqlite3.connect(db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        conn.row_factory = sqlite3.Row
        
        # Check if migration is needed
        old_data = get_old_schema_data(conn)
        
        if old_data is None:
            print("\nNo migration needed. Database already uses new schema.")
            conn.close()
            return
        
        # Extract normalized data
        print("\n→ Extracting and normalizing data...")
        companies = extract_companies(old_data)
        products = extract_products(old_data)
        
        # Try to load NIP data from company_db if available
        try:
            import sys
            sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'ExcelVerifier'))
            from core.company_db import load_company_db
            company_db = load_company_db()
            
            # Match companies with their NIPs
            for company in companies:
                nip = company_db.get(company['name'])
                if nip:
                    company['nip'] = nip
            
            print(f"✓ Loaded NIP data for {sum(1 for c in companies if c['nip'])} companies")
        except Exception as e:
            print(f"  ⚠ Could not load company_db.py: {e}")
        
        # Create new schema
        print("\n→ Creating new database structure...")
        create_new_schema(conn)
        
        # Populate lookup tables
        print("\n→ Populating normalized tables...")
        company_map = populate_companies(conn, companies)
        product_map = populate_products(conn, products)
        
        # Migrate transactional data
        print("\n→ Migrating transactional data...")
        migrate_data(conn, old_data, company_map, product_map)
        
        # Cleanup
        print("\n→ Cleaning up...")
        cleanup_old_tables(conn)
        
        conn.close()
        
        print("\n" + "="*60)
        print("✓ MIGRATION COMPLETED SUCCESSFULLY!")
        print("="*60)
        print(f"\nBackup saved at: {backup_path}")
        print("You can delete the backup file once you've verified the migration.\n")
        
    except Exception as e:
        print(f"\n✗ MIGRATION FAILED: {e}")
        print(f"Your original database is backed up at: {backup_path}")
        print("You can restore it by copying the backup file back to excelverifier.db\n")
        raise


if __name__ == "__main__":
    main()
