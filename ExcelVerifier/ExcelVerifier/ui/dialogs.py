import sys
import os
import re
from datetime import datetime

# --- PATH FIX: FORCE PYTHON TO FIND CONFIG.PY ---
# This ensures imports work regardless of which folder you run the script from.
current_dir = os.path.dirname(os.path.abspath(__file__)) # .../ui
parent_dir = os.path.dirname(current_dir)                # .../ExcelVerifier (Root)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)
# -----------------------------------------------

import config 
import pandas as pd
from core.company_db import load_company_db, normalize_nip

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QAbstractItemView, QCalendarWidget, QToolButton, QMenu, QWidgetAction,
    QFrame, QWidget
)
from PyQt5.QtCore import Qt, QLocale
from PyQt5.QtGui import QIcon, QColor, QBrush

# --- Custom Calendar Button (Polish + White + Select Month) ---
class DatePickerButton(QToolButton):
    def __init__(self, target_line_edit):
        super().__init__()
        self.setText("📅") 
        self.setToolTip("Wybierz datę") 
        self.target = target_line_edit
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedSize(40, 38) 
        
        # Style the button to merge with input
        self.setStyleSheet("""
            QToolButton {
                border: 1px solid #D1D5DB;
                border-left: none;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
                background-color: #F3F4F6;
            }
            QToolButton:hover { background-color: #E5E7EB; }
        """)
        
        # Container for Calendar + Bottom Button
        self.container = QWidget()
        self.container_layout = QVBoxLayout(self.container)
        self.container_layout.setContentsMargins(0, 0, 0, 0)
        self.container_layout.setSpacing(0)

        # 1. The Calendar
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)
        # Set Polish Locale
        self.calendar.setLocale(QLocale(QLocale.Polish))
        self.calendar.clicked.connect(self.on_date_selected)
        
        # White/Modern Styling
        self.calendar.setStyleSheet("""
            QCalendarWidget QWidget { background-color: white; color: #1F2937; }
            QCalendarWidget QTableView {
                background-color: white;
                alternate-background-color: #F9FAFB;
                selection-background-color: #2563EB; 
                selection-color: white;
            }
            QCalendarWidget QToolButton {
                color: #1F2937;
                background-color: transparent;
                icon-size: 14px;
                font-weight: bold;
            }
            QCalendarWidget QToolButton:hover {
                background-color: #E5E7EB;
                border-radius: 4px;
            }
            QCalendarWidget QMenu { background-color: white; color: #1F2937; }
            QCalendarWidget QSpinBox { background-color: white; color: #1F2937; }
        """)

        # 2. "Select Whole Month" Button
        self.month_btn = QPushButton("Wybierz cały miesiąc")
        self.month_btn.setCursor(Qt.PointingHandCursor)
        self.month_btn.clicked.connect(self.on_month_selected)
        self.month_btn.setStyleSheet("""
            QPushButton {
                background-color: #F3F4F6;
                border: none;
                border-top: 1px solid #E5E7EB;
                color: #374151;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #E5E7EB;
                color: #2563EB;
            }
        """)

        self.container_layout.addWidget(self.calendar)
        self.container_layout.addWidget(self.month_btn)

        # 3. Popup Menu
        self.menu = QMenu(self)
        action = QWidgetAction(self.menu)
        action.setDefaultWidget(self.container)
        self.menu.addAction(action)
        self.setMenu(self.menu)
        self.setPopupMode(QToolButton.InstantPopup)

    def on_date_selected(self, qdate):
        # User picked a specific day
        self.target.setText(qdate.toString("yyyy-MM-dd"))
        self.menu.close()

    def on_month_selected(self):
        # User picked the month
        curr_month = self.calendar.monthShown()
        curr_year = self.calendar.yearShown()
        date_str = f"{curr_year}-{curr_month:02d}"
        self.target.setText(date_str)
        self.menu.close()

# --- Main Approved Reports Window ---
class ApprovedReportsDialog(QDialog):
    def __init__(self, parent=None, filter_month=None):
        super().__init__(parent)
        self.setWindowTitle("Zatwierdzone raporty")
        self.resize(950, 650)
        self.selected_file_path = None
        self.df = pd.DataFrame()
        self.df_filtered_by_month = pd.DataFrame()  # Store month-filtered data
        self.filter_month = filter_month  # YYYY-MM format
        
        if parent:
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.load_data()

    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)

        # Title
        title_lbl = QLabel("Zatwierdzone raporty")
        title_lbl.setStyleSheet("font-size: 20px; font-weight: bold; color: #111827;")
        layout.addWidget(title_lbl)

        # Filter Card
        filter_frame = QFrame()
        filter_frame.setObjectName("FilterFrame")
        filter_frame.setStyleSheet("""
            QFrame#FilterFrame {
                background-color: white;
                border: 1px solid #E5E7EB;
                border-radius: 8px;
            }
            QLineEdit {
                border: 1px solid #D1D5DB;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 14px;
                background-color: #FFFFFF;
                color: #374151;
            }
            QLineEdit:focus { border: 1px solid #2563EB; }
        """)
        
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(15, 15, 15, 15)
        filter_layout.setSpacing(12)

        # Company Input
        self.company_filter = QLineEdit()
        self.company_filter.setPlaceholderText("Szukaj po firmie...")
        self.company_filter.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.company_filter, 1)

        # NIP Input
        self.nip_filter = QLineEdit()
        self.nip_filter.setPlaceholderText("Szukaj po NIP...")
        self.nip_filter.setFixedWidth(140)
        self.nip_filter.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.nip_filter, 0)

        # Date Input + Button
        date_container = QWidget()
        date_container.setStyleSheet("background: transparent; border: none;") 
        date_box = QHBoxLayout(date_container)
        date_box.setContentsMargins(0, 0, 0, 0)
        date_box.setSpacing(0)

        self.date_filter = QLineEdit()
        self.date_filter.setPlaceholderText("RRRR-MM-DD")
        self.date_filter.setFixedWidth(120)
        self.date_filter.setStyleSheet("border-top-right-radius: 0px; border-bottom-right-radius: 0px;")
        self.date_filter.textChanged.connect(self.apply_filters)
        
        self.date_btn = DatePickerButton(self.date_filter)

        date_box.addWidget(self.date_filter)
        date_box.addWidget(self.date_btn)
        
        filter_layout.addWidget(date_container, 0)
        layout.addWidget(filter_frame)

        # Table
        self.table = QTableWidget()
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #E5E7EB;
                border-radius: 6px;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #F9FAFB;
                padding: 10px;
                border: none;
                border-bottom: 2px solid #E5E7EB;
                font-weight: 600;
                color: #374151;
            }
        """)
        layout.addWidget(self.table)

        # Footer Buttons
        btn_layout = QHBoxLayout()
        
        self.close_btn = QPushButton("Zamknij")
        self.close_btn.setFixedSize(100, 45)
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.clicked.connect(self.close_dialog) 

        self.open_btn = QPushButton("Otwórz wybrany raport")
        self.open_btn.setObjectName("PrimaryBtn")
        self.open_btn.setFixedSize(220, 45)
        self.open_btn.setCursor(Qt.PointingHandCursor)
        self.open_btn.clicked.connect(self.accept_selection)
        
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.close_btn)
        btn_layout.addWidget(self.open_btn)
        
        layout.addLayout(btn_layout)

    def load_data(self):
        try:
            from core.database_handler import DatabaseHandler
            from config import DATABASE_FILE
            
            db = DatabaseHandler(DATABASE_FILE)
            records = db.get_all_approved_records()
            
            # Convert to DataFrame
            self.df = pd.DataFrame(records)
            
            # Rename columns to match expected format (new schema uses company_name, company_nip)
            if not self.df.empty:
                self.df = self.df.rename(columns={
                    'date': 'Date',
                    'company_name': 'Company',
                    'company_nip': 'NIP',
                    'filename': 'Filename',
                    'filepath': 'Filepath'
                })

                # Normalize NIP column (new schema has direct NIP from companies table)
                if 'NIP' in self.df.columns:
                    self.df['NIP'] = self.df['NIP'].apply(lambda x: normalize_nip(x) if x else "")
                else:
                    self.df['NIP'] = ""
                
                # Sort by date descending
                self.df['_sort_date'] = pd.to_datetime(self.df['Date'], errors='coerce')
                self.df = self.df.sort_values(by='_sort_date', ascending=False)
                self.df.drop(columns=['_sort_date'], inplace=True)
                
                # Filter by month if provided (for navigation only)
                if self.filter_month:
                    self.df_filtered_by_month = self.df[self.df['Date'].astype(str).str.startswith(self.filter_month)]
                else:
                    # Get current month for navigation
                    current_month = datetime.now().strftime('%Y-%m')
                    self.df_filtered_by_month = self.df[self.df['Date'].astype(str).str.startswith(current_month)]
                    
                    # If no records for current month, use the latest month available
                    if self.df_filtered_by_month.empty:
                        latest_date = self.df['Date'].astype(str).max()
                        if latest_date:
                            latest_month = latest_date[:7]
                            self.df_filtered_by_month = self.df[self.df['Date'].astype(str).str.startswith(latest_month)]
                
                # If still empty, use all for navigation
                if self.df_filtered_by_month.empty:
                    self.df_filtered_by_month = self.df.copy()
                
                # Display ALL reports in table
                self.populate_table(self.df)
            else:
                self.df = pd.DataFrame(columns=["Date", "Company", "Filename", "Filepath"])
                self.df_filtered_by_month = self.df.copy()
                self.populate_table(self.df)
                
        except Exception as e:
            print(f"Error loading approved records: {e}")
            self.df = pd.DataFrame(columns=["Date", "Company", "Filename", "Filepath"])
            self.df_filtered_by_month = self.df.copy()
            self.populate_table(self.df_filtered_by_month)

    def populate_table(self, df_to_show):
        self.table.clear()
        cols = ["Data", "Firma", "Nazwa pliku"]
        self.table.setColumnCount(4) # +1 Hidden
        self.table.setHorizontalHeaderLabels(cols + [""])
        self.table.setRowCount(len(df_to_show))

        for r, row in df_to_show.reset_index(drop=True).iterrows():
            val_date = str(row.get('Date', '')).split(" ")[0]
            filepath = str(row.get('Filepath', ''))
            print(f"[DEBUG] Populating row {r}: filepath={filepath}")
            self.table.setItem(r, 0, QTableWidgetItem(val_date))
            self.table.setItem(r, 1, QTableWidgetItem(str(row.get('Company', ''))))
            self.table.setItem(r, 2, QTableWidgetItem(str(row.get('Filename', ''))))
            self.table.setItem(r, 3, QTableWidgetItem(filepath))

        self.table.setColumnHidden(3, True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)

    def apply_filters(self):
        if self.df.empty:
            return
        
        comp_txt = self.company_filter.text().lower().strip()
        nip_txt = self.nip_filter.text().strip()
        date_txt = self.date_filter.text().lower().strip()
        
        # Start with all records
        filtered = self.df.copy()
        
        if comp_txt:
            filtered = filtered[filtered['Company'].astype(str).str.lower().str.contains(comp_txt)]
        if nip_txt:
            nip_search = re.sub(r"\D", "", nip_txt)
            nip_letters = re.sub(r"\d", " ", nip_txt).lower().strip()
            nip_letters = " ".join(nip_letters.split())
            
            masks = []
            if nip_search:
                nip_mask = filtered['NIP'].astype(str).str.contains(nip_search)
                company_digit_mask = filtered['Company'].astype(str).apply(
                    lambda v: nip_search in re.sub(r"\D", "", v)
                )
                masks.append(nip_mask | company_digit_mask)
            if nip_letters:
                company_text_mask = filtered['Company'].astype(str).str.lower().str.contains(nip_letters)
                masks.append(company_text_mask)
            if masks:
                combined = masks[0]
                for m in masks[1:]:
                    combined = combined & m
                filtered = filtered[combined]
        if date_txt:
            filtered = filtered[filtered['Date'].astype(str).str.contains(date_txt)]
        
        self.populate_table(filtered)

    def _extract_nip(self, text):
        if not text:
            return ""
        value = str(text).strip()
        normalized = normalize_nip(value)
        if len(normalized) == 10:
            return normalized
        # Try formatted patterns first
        match = re.search(r'(\d{3})-(\d{2})-(\d{2})-(\d{3})', value)
        if match:
            return "".join(match.groups())
        match = re.search(r'(\d{3})-(\d{3})-(\d{2})-(\d{2})', value)
        if match:
            return "".join(match.groups())
        # Try 10 consecutive digits
        match = re.search(r'(?:^|\s|[^\d])(\d{10})(?:\s|$|[^\d])', value)
        if match:
            return match.group(1)
        return ""

    def _normalize_company_name(self, value):
        text = str(value).strip().lower()
        # Remove embedded NIP patterns to improve name matching
        text = re.sub(r"\bnip\b[:.\s]*\d{10}", " ", text)
        text = re.sub(r"\d{10}", " ", text)
        return " ".join(text.split())

    def accept_selection(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Brak wyboru", "Proszę wybrać raport.")
            return
        path_item = self.table.item(row, 3)
        if path_item:
            self.selected_file_path = path_item.text()
            if not os.path.exists(self.selected_file_path):
                 QMessageBox.warning(self, "Brak", "Nie znaleziono pliku.")
                 return
            self.accept()
        else:
            QMessageBox.warning(self, "Błąd", "Nie można odczytać ścieżki pliku z tabeli.")
    
    
    def close_dialog(self):
        """Close the dialog without selecting anything."""
        self.selected_file_path = None
        self.reject()
    
    def closeEvent(self, event):
        """Override close event to ensure dialog closes without affecting parent."""
        event.accept()


# --- Unapproved Reports Dialog ---
class UnapprovedReportsDialog(QDialog):
    def __init__(self, unapproved_list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Niezatwierdzony raport")
        self.resize(800, 500)
        self.selected_file_path = None
        self.unapproved_list = unapproved_list
        
        if parent:
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.populate_table()

    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)

        # Title
        title_lbl = QLabel("Niezatwierdzony raport")
        title_lbl.setStyleSheet("font-size: 20px; font-weight: bold; color: #111827;")
        layout.addWidget(title_lbl)

        # Table
        self.table = QTableWidget()
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #E5E7EB;
                border-radius: 6px;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #F9FAFB;
                padding: 10px;
                border: none;
                border-bottom: 2px solid #E5E7EB;
                font-weight: 600;
                color: #374151;
            }
        """)
        layout.addWidget(self.table)

        # Footer Buttons
        btn_layout = QHBoxLayout()
        
        self.close_btn = QPushButton("Zamknij")
        self.close_btn.setFixedSize(100, 45)
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.clicked.connect(self.close_dialog)

        self.open_btn = QPushButton("Otwórz raport")
        self.open_btn.setObjectName("PrimaryBtn")
        self.open_btn.setFixedSize(220, 45)
        self.open_btn.setCursor(Qt.PointingHandCursor)
        self.open_btn.clicked.connect(self.accept_selection)
        
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.close_btn)
        btn_layout.addWidget(self.open_btn)
        
        layout.addLayout(btn_layout)

    def populate_table(self):
        self.table.clear()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["#", "Nazwa pliku"])
        self.table.setRowCount(len(self.unapproved_list))

        for i, path in enumerate(self.unapproved_list):
            self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.table.setItem(i, 1, QTableWidgetItem(os.path.basename(path)))

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)

    def accept_selection(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Brak wyboru", "Proszę wybrać raport.")
            return
        self.selected_file_path = self.unapproved_list[row]
        self.accept()
    
    def close_dialog(self):
        """Close the dialog without selecting anything."""
        self.selected_file_path = None
        self.reject()
    
    def closeEvent(self, event):
        """Override close event to ensure dialog closes without affecting parent."""
        event.accept()