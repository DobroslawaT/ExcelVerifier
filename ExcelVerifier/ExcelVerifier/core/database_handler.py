"""
Database handler for ExcelVerifier application.
Manages SQLite database for approved records and reporting data.
"""

import sqlite3
import os
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from contextlib import contextmanager


class DatabaseHandler:
    """Handles all database operations for the application."""
    
    def __init__(self, db_path: str = "excelverifier.db"):
        """
        Initialize database handler.
        
        Args:
            db_path: Path to SQLite database file
        """
        self.db_path = db_path
        self._initialize_database()
    
    @contextmanager
    def _get_connection(self):
        """Context manager for database connections."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row  # Enable column access by name
        conn.execute("PRAGMA foreign_keys = ON")  # Enable foreign key constraints
        try:
            yield conn
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def _initialize_database(self):
        """Create tables if they don't exist with proper foreign key relationships."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            
            # Create companies table (no dependencies)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS companies (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    nip TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_companies_name
                ON companies(name)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_companies_nip
                ON companies(nip)
            """)
            
            # Create products table (no dependencies)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    code TEXT
                )
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_products_name
                ON products(name)
            """)
            
            # Create orders table (depends on companies)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS orders (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    company_id INTEGER NOT NULL,
                    date_issued TEXT NOT NULL,
                    document_number TEXT,
                    FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE
                )
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_orders_company
                ON orders(company_id)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_orders_date
                ON orders(date_issued)
            """)
            
            # Create approved_records table (depends on orders)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS approved_records (
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
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_approved_date 
                ON approved_records(date)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_approved_filename 
                ON approved_records(filename)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_approved_order
                ON approved_records(order_id)
            """)
            
            # Create order_items table (depends on orders and products)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS order_items (
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
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_order_items_order
                ON order_items(order_id)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_order_items_product
                ON order_items(product_id)
            """)
            
            conn.commit()
    
    # ==================== COMPANIES ====================
    
    def add_company(self, name: str, nip: str = None) -> Optional[int]:
        """
        Add a new company or get existing company ID.
        
        Args:
            name: Company name (unique)
            nip: Optional NIP number
            
        Returns:
            Company ID
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute("""
                    INSERT INTO companies (name, nip)
                    VALUES (?, ?)
                """, (name, nip))
                return cursor.lastrowid
            except sqlite3.IntegrityError:
                # Company already exists, get its ID
                cursor.execute("SELECT id FROM companies WHERE name = ?", (name,))
                row = cursor.fetchone()
                return row['id'] if row else None
    
    def get_company_by_id(self, company_id: int) -> Optional[Dict]:
        """Get company by ID."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, name, nip, created_at, updated_at
                FROM companies WHERE id = ?
            """, (company_id,))
            row = cursor.fetchone()
            return dict(row) if row else None
    
    def get_company_by_name(self, name: str) -> Optional[Dict]:
        """Get company by name."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, name, nip, created_at, updated_at
                FROM companies WHERE name = ?
            """, (name,))
            row = cursor.fetchone()
            return dict(row) if row else None
    
    def get_all_companies(self) -> List[Dict]:
        """Get all companies."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, name, nip, created_at, updated_at
                FROM companies
                ORDER BY name ASC
            """)
            return [dict(row) for row in cursor.fetchall()]
    
    def update_company(self, company_id: int, name: str = None, nip: str = None) -> bool:
        """Update company details."""
        updates = []
        params = []
        
        if name is not None:
            updates.append("name = ?")
            params.append(name)
        if nip is not None:
            updates.append("nip = ?")
            params.append(nip)
        
        if not updates:
            return False
        
        updates.append("updated_at = CURRENT_TIMESTAMP")
        params.append(company_id)
        
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"""
                UPDATE companies 
                SET {', '.join(updates)}
                WHERE id = ?
            """, params)
            return cursor.rowcount > 0
    
    def delete_company(self, company_id: int) -> bool:
        """Delete a company (cascades to orders and order_items)."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM companies WHERE id = ?", (company_id,))
            return cursor.rowcount > 0
    
    # ==================== PRODUCTS ====================
    
    def add_product(self, name: str, code: str = None) -> int:
        """
        Add a new product or get existing product ID.
        
        Args:
            name: Product name
            code: Optional product code
            
        Returns:
            Product ID
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute("""
                    INSERT INTO products (name, code)
                    VALUES (?, ?)
                """, (name, code))
                return cursor.lastrowid
            except sqlite3.IntegrityError:
                # Product already exists, get its ID
                cursor.execute("""
                    SELECT id FROM products 
                    WHERE name = ?
                """, (name,))
                row = cursor.fetchone()
                return row['id'] if row else None
    
    def get_product_by_id(self, product_id: int) -> Optional[Dict]:
        """Get product by ID."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, name, code
                FROM products WHERE id = ?
            """, (product_id,))
            row = cursor.fetchone()
            return dict(row) if row else None
    
    def get_all_products(self) -> List[Dict]:
        """Get all products."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, name, code
                FROM products
                ORDER BY name ASC
            """)
            return [dict(row) for row in cursor.fetchall()]
    
    # ==================== ORDERS ====================
    
    def add_order(self, company_id: int, date_issued: str, document_number: str = None) -> int:
        """
        Add a new order.
        
        Args:
            company_id: Foreign key to companies table
            date_issued: Date in yyyy-MM-dd format
            document_number: Optional document/order number
            
        Returns:
            Order ID
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO orders (company_id, date_issued, document_number)
                VALUES (?, ?, ?)
            """, (company_id, date_issued, document_number))
            return cursor.lastrowid
    
    def get_order_by_id(self, order_id: int) -> Optional[Dict]:
        """Get order by ID with company details."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT o.id, o.company_id, o.date_issued, o.document_number,
                       c.name as company_name, c.nip as company_nip
                FROM orders o
                JOIN companies c ON o.company_id = c.id
                WHERE o.id = ?
            """, (order_id,))
            row = cursor.fetchone()
            return dict(row) if row else None
    
    def get_orders_by_company(self, company_id: int) -> List[Dict]:
        """Get all orders for a company."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, company_id, date_issued, document_number
                FROM orders
                WHERE company_id = ?
                ORDER BY date_issued DESC
            """, (company_id,))
            return [dict(row) for row in cursor.fetchall()]
    
    def get_all_orders(self) -> List[Dict]:
        """Get all orders with company details."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT o.id, o.company_id, o.date_issued, o.document_number,
                       c.name as company_name, c.nip as company_nip
                FROM orders o
                JOIN companies c ON o.company_id = c.id
                ORDER BY o.date_issued DESC
            """)
            return [dict(row) for row in cursor.fetchall()]
    
    # ==================== APPROVED RECORDS ====================
    
    def add_approved_record(self, order_id: int, date: str, filename: str, filepath: str) -> Optional[int]:
        """
        Add a new approved record.
        
        Args:
            order_id: Foreign key to orders table
            date: Date in yyyy-MM-dd format
            filename: Excel filename (unique)
            filepath: Full path to Excel file
            
        Returns:
            Record ID or None if duplicate filename
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO approved_records (order_id, date, filename, filepath)
                    VALUES (?, ?, ?, ?)
                """, (order_id, date, filename, filepath))
                return cursor.lastrowid
        except sqlite3.IntegrityError:
            # Duplicate filename
            return None
    
    def update_approved_date(self, filename: str, new_date: str) -> bool:
        """
        Update the date for an approved record.
        
        Args:
            filename: Excel filename
            new_date: New date in yyyy-MM-dd format
            
        Returns:
            True if updated, False if not found
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE approved_records 
                SET date = ?, updated_at = CURRENT_TIMESTAMP
                WHERE filename = ?
            """, (new_date, filename))
            return cursor.rowcount > 0
    
    def delete_approved_record(self, filename: str) -> bool:
        """
        Delete an approved record.
        
        Args:
            filename: Excel filename
            
        Returns:
            True if deleted, False if not found
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM approved_records WHERE filename = ?", (filename,))
            return cursor.rowcount > 0

    def delete_reporting_data_by_filename(self, filename: str) -> bool:
        """
        Delete reporting data (orders and order_items) for a given approved filename.

        Args:
            filename: Excel filename

        Returns:
            True if related data was deleted, False if not found
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT order_id FROM approved_records WHERE filename = ?", (filename,))
            row = cursor.fetchone()
            if not row:
                return False

            order_id = row["order_id"]
            cursor.execute("DELETE FROM order_items WHERE order_id = ?", (order_id,))
            cursor.execute("DELETE FROM orders WHERE id = ?", (order_id,))
            return True
    
    def get_approved_record(self, filename: str) -> Optional[Dict]:
        """
        Get a specific approved record with order and company details.
        
        Args:
            filename: Excel filename
            
        Returns:
            Dictionary with record data or None if not found
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT ar.id, ar.order_id, ar.date, ar.filename, ar.filepath, 
                       ar.created_at, ar.updated_at,
                       o.company_id, o.date_issued, o.document_number,
                       c.name as company_name, c.nip as company_nip
                FROM approved_records ar
                JOIN orders o ON ar.order_id = o.id
                JOIN companies c ON o.company_id = c.id
                WHERE ar.filename = ?
            """, (filename,))
            row = cursor.fetchone()
            return dict(row) if row else None
    
    def get_all_approved_records(self) -> List[Dict]:
        """
        Get all approved records with order and company details.
        
        Returns:
            List of dictionaries with record data
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT ar.id, ar.order_id, ar.date, ar.filename, ar.filepath,
                       ar.created_at, ar.updated_at,
                       o.company_id, o.date_issued, o.document_number,
                       c.name as company_name, c.nip as company_nip
                FROM approved_records ar
                JOIN orders o ON ar.order_id = o.id
                JOIN companies c ON o.company_id = c.id
                ORDER BY ar.date DESC, c.name ASC
            """)
            return [dict(row) for row in cursor.fetchall()]
    
    def get_available_months(self) -> List[str]:
        """
        Get list of unique months from approved records.
        
        Returns:
            List of month strings in YYYY-MM format, sorted descending
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT MIN(date) as min_date
                FROM approved_records
                WHERE date IS NOT NULL AND date != ''
            """)
            row = cursor.fetchone()
            if not row or not row["min_date"]:
                return []

            try:
                start_date = datetime.strptime(row["min_date"], "%Y-%m-%d")
            except ValueError:
                return []

            today = datetime.now()
            start_year = start_date.year
            start_month = start_date.month
            end_year = today.year
            end_month = today.month

            months = []
            year, month = start_year, start_month
            while (year < end_year) or (year == end_year and month <= end_month):
                months.append(f"{year:04d}-{month:02d}")
                month += 1
                if month > 12:
                    month = 1
                    year += 1

            months.reverse()
            return months
    
    def get_approved_records_by_month(self, year_month: str) -> List[Dict]:
        """
        Get approved records for a specific month with company details.
        
        Args:
            year_month: Month in YYYY-MM format (e.g., "2026-02")
            
        Returns:
            List of dictionaries with record data
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT ar.id, ar.order_id, ar.date, ar.filename, ar.filepath,
                       ar.created_at, ar.updated_at,
                       o.company_id, o.date_issued, o.document_number,
                       c.name as company_name, c.nip as company_nip
                FROM approved_records ar
                JOIN orders o ON ar.order_id = o.id
                JOIN companies c ON o.company_id = c.id
                WHERE ar.date LIKE ?
                ORDER BY ar.date ASC, c.name ASC
            """, (f"{year_month}%",))
            return [dict(row) for row in cursor.fetchall()]
    
    def get_approved_records_filtered(self, month: Optional[str] = None, company_name: Optional[str] = None) -> List[Dict]:
        """
        Get approved records with optional filters.
        
        Args:
            month: Filter by month (YYYY-MM format)
            company_name: Filter by company name
        
        Returns:
            List of dictionaries, sorted by date descending
        """
        query = """
            SELECT ar.id, ar.order_id, ar.date, ar.filename, ar.filepath,
                   ar.created_at, ar.updated_at,
                   o.company_id, o.date_issued, o.document_number,
                   c.name as company_name, c.nip as company_nip
            FROM approved_records ar
            JOIN orders o ON ar.order_id = o.id
            JOIN companies c ON o.company_id = c.id
            WHERE 1=1
        """
        params = []
        
        if month:
            query += " AND ar.date LIKE ?"
            params.append(f"{month}%")
        
        if company_name:
            query += " AND c.name = ?"
            params.append(company_name)
        
        query += " ORDER BY ar.date DESC"
        
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
    
    def get_approved_companies(self) -> List[str]:
        """
        Get list of unique companies from approved records.
        
        Returns:
            Sorted list of company names
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT DISTINCT c.name
                FROM approved_records ar
                JOIN orders o ON ar.order_id = o.id
                JOIN companies c ON o.company_id = c.id
                ORDER BY c.name ASC
            """)
            return [row['name'] for row in cursor.fetchall()]
    
    # ==================== ORDER ITEMS ====================
    
    def add_order_items(self, items: List[Dict]) -> int:
        """
        Add multiple order items.
        
        Args:
            items: List of dictionaries with keys:
                - order_id: Foreign key to orders
                - product_id: Foreign key to products
                - quantity_delivery: Delivery quantity
                - quantity_return: Return quantity
                - previous_state: Previous state/stock
                - state_after: State after transaction
        
        Returns:
            Number of items inserted
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany("""
                INSERT INTO order_items 
                (order_id, product_id, quantity_delivery, quantity_return, 
                 previous_state, state_after)
                VALUES (:order_id, :product_id, :quantity_delivery, :quantity_return,
                        :previous_state, :state_after)
            """, items)
            return cursor.rowcount
    
    def get_order_items(self, order_id: int) -> List[Dict]:
        """
        Get all items for an order.
        
        Args:
            order_id: Order ID
            
        Returns:
            List of order items with product details
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT oi.id, oi.order_id, oi.product_id,
                       oi.quantity_delivery, oi.quantity_return,
                       oi.previous_state, oi.state_after, oi.created_at,
                       p.name as product_name, p.code as product_code
                FROM order_items oi
                JOIN products p ON oi.product_id = p.id
                WHERE oi.order_id = ?
                ORDER BY p.name ASC
            """, (order_id,))
            return [dict(row) for row in cursor.fetchall()]
    
    def get_all_order_items_with_details(self, filters: Optional[Dict] = None) -> List[Dict]:
        """
        Get all order items with full order, company, and product details.
        
        Args:
            filters: Optional dictionary with:
                - month: YYYY-MM format
                - company_name: Company name
                - start_date: yyyy-MM-dd
                - end_date: yyyy-MM-dd
        
        Returns:
            List of dictionaries with complete item data
        """
        query = """
            SELECT oi.id, oi.order_id, oi.product_id,
                   oi.quantity_delivery, oi.quantity_return,
                   oi.previous_state, oi.state_after, oi.created_at,
                   o.date_issued, o.document_number,
                   c.id as company_id, c.name as company_name, c.nip as company_nip,
                   p.name as product_name, p.code as product_code,
                   ar.filename, ar.filepath
            FROM order_items oi
            JOIN orders o ON oi.order_id = o.id
            JOIN companies c ON o.company_id = c.id
            JOIN products p ON oi.product_id = p.id
            LEFT JOIN approved_records ar ON ar.order_id = o.id
            WHERE 1=1
        """
        params = []
        
        if filters:
            if filters.get('month'):
                query += " AND o.date_issued LIKE ?"
                params.append(f"{filters['month']}%")
            
            if filters.get('company_name'):
                query += " AND c.name = ?"
                params.append(filters['company_name'])
            
            if filters.get('start_date'):
                query += " AND o.date_issued >= ?"
                params.append(filters['start_date'])
            
            if filters.get('end_date'):
                query += " AND o.date_issued <= ?"
                params.append(filters['end_date'])
        
        query += " ORDER BY o.date_issued DESC, c.name ASC, p.name ASC"
        
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
    
    # ==================== UTILITY METHODS ====================
    
    def get_database_stats(self) -> Dict:
        """
        Get database statistics.
        
        Returns:
            Dictionary with counts and info
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            
            cursor.execute("SELECT COUNT(*) as count FROM companies")
            companies_count = cursor.fetchone()['count']
            
            cursor.execute("SELECT COUNT(*) as count FROM products")
            products_count = cursor.fetchone()['count']
            
            cursor.execute("SELECT COUNT(*) as count FROM orders")
            orders_count = cursor.fetchone()['count']
            
            cursor.execute("SELECT COUNT(*) as count FROM approved_records")
            approved_count = cursor.fetchone()['count']
            
            cursor.execute("SELECT COUNT(*) as count FROM order_items")
            items_count = cursor.fetchone()['count']
            
            cursor.execute("SELECT MIN(date) as earliest, MAX(date) as latest FROM approved_records")
            date_range = cursor.fetchone()
            
            return {
                'companies_count': companies_count,
                'products_count': products_count,
                'orders_count': orders_count,
                'approved_records_count': approved_count,
                'order_items_count': items_count,
                'earliest_date': date_range['earliest'],
                'latest_date': date_range['latest'],
                'database_size': os.path.getsize(self.db_path) if os.path.exists(self.db_path) else 0
            }
    
    def vacuum_database(self):
        """Optimize database by reclaiming unused space."""
        with self._get_connection() as conn:
            conn.execute("VACUUM")
    
    # ==================== LEGACY SUPPORT FOR COMPATIBILITY ====================
    
    def upsert_company(self, name: str, nip: str) -> None:
        """Legacy method: Insert or update a company."""
        self.add_company(name=name, nip=nip)
    
    def replace_companies(self, companies: List[Dict]) -> None:
        """Legacy method: Replace all companies with provided list."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM companies")
            for company in companies:
                name = company.get('name', '')
                nip = company.get('nip', '')
                if name:  # Only add if name exists
                    try:
                        cursor.execute("""
                            INSERT INTO companies (name, nip)
                            VALUES (?, ?)
                        """, (name, nip))
                    except sqlite3.IntegrityError:
                        # Company already exists with this name, update it
                        cursor.execute("""
                            UPDATE companies SET nip = ?, updated_at = CURRENT_TIMESTAMP
                            WHERE name = ?
                        """, (nip, name))
    
    def delete_company_by_nip(self, nip: str) -> bool:
        """Legacy method: Delete a company by NIP."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM companies WHERE nip = ?", (nip,))
            row = cursor.fetchone()
            if row:
                return self.delete_company(row['id'])
            return False
    
    def get_companies(self) -> List[Dict]:
        """Legacy method: Get all companies."""
        return self.get_all_companies()
