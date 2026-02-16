from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QComboBox, QDateEdit, QMessageBox, QProgressDialog, QCheckBox, QSizePolicy, QFileDialog
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal, QThread, QSize
from PyQt5.QtGui import QFont
from core.excel_handler import ExcelHandler
from core.file_manager import FileManager
import os


class ReportWorker(QThread):
    """Worker thread for generating reports."""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, excel_handler, file_manager, filters, output_path=None):
        super().__init__()
        self.excel_handler = excel_handler
        self.file_manager = file_manager
        self.filters = filters
        self.output_path = output_path
    
    def run(self):
        try:
            self.progress.emit("Rozpoczynanie generowania raportu...")
            result = self.excel_handler.generate_report(self.filters, self.output_path)
            self.finished.emit(True, f"Raport wygenerowany pomyślnie:\n{result}")
        except Exception as e:
            self.finished.emit(False, f"Błąd generowania raportu:\n{str(e)}")


class GenerateReportPage(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_handler = ExcelHandler()
        self.file_manager = FileManager()
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        # Title
        title = QLabel("Generuj raport")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #111827; margin-bottom: 10px;")
        title.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(title)
        
        # Description
        desc = QLabel("Wybierz miesiąc lub wszystkie miesiące i generuj raport z zatwierdzonych plików Excel.")
        desc.setStyleSheet("font-size: 14px; color: #6B7280; margin-bottom: 20px;")
        desc.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(desc)
        
        # Month Selection
        month_label = QLabel("Wybierz miesiąc:")
        month_label.setStyleSheet("font-size: 13px; font-weight: 600; color: #374151;")
        month_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(month_label)
        
        # Month combo box
        self.month_combo = QComboBox()
        self.month_combo.addItems(self.get_available_months())
        self.month_combo.setStyleSheet("""
            QComboBox {
                padding: 8px 12px;
                border: 1px solid #D1D5DB;
                border-radius: 6px;
                font-size: 13px;
                background-color: white;
            }
            QComboBox:focus { border: 2px solid #2563EB; }
        """)
        self.month_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(self.month_combo)
        
        # Stretch to push button to bottom
        layout.addStretch(1)
        
        # Generate Button - ensure it stays at bottom with proper sizing
        button_widget = QWidget()
        button_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        button_widget.setMinimumHeight(80)
        button_container = QHBoxLayout(button_widget)
        button_container.setContentsMargins(0, 0, 0, 0)
        button_container.setSpacing(10)
        button_container.addStretch()
        
        self.generate_btn = QPushButton("📊 Generuj Raport")
        self.generate_btn.setFixedSize(200, 50)
        self.generate_btn.setCursor(Qt.PointingHandCursor)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #2563EB;
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #1D4ED8; }
        """)
        self.generate_btn.clicked.connect(self.generate_report)
        button_container.addWidget(self.generate_btn)
        button_container.addStretch()
        
        layout.addWidget(button_widget)
    
    def refresh_months(self):
        """Refresh the month combo box with latest data."""
        self.month_combo.clear()
        self.month_combo.addItems(self.get_available_months())
    
    def get_available_months(self):
        """Get list of available months from database (fast!)."""
        try:
            from core.database_handler import DatabaseHandler
            from config import DATABASE_FILE
            
            db = DatabaseHandler(DATABASE_FILE)
            months = db.get_available_months()
            
            if months:
                return ["Wszystkie miesiące"] + months
            else:
                return ["Nie znaleziono miesięcy"]
        except Exception as e:
            print(f"Error reading database: {e}")
            return ["Błąd ładowania miesięcy"]
    
    def generate_report(self):
        """Generate report for selected month."""
        selected_month = self.month_combo.currentText()
        
        if selected_month in ["Nie znaleziono miesięcy", "Błąd ładowania miesięcy"]:
            QMessageBox.warning(self, "Ostrzeżenie", "Nie wybrano prawidłowego miesiąca.")
            return

        # Build filters dictionary
        if selected_month == "Wszystkie miesiące":
            filters = {
                'mode': 0,  # All months
                'month': None,
                'from_date': None,
                'to_date': None,
                'company': None
            }
        else:
            filters = {
                'mode': 1,  # Month filter mode
                'month': selected_month,
                'from_date': None,
                'to_date': None,
                'company': None
            }
        
        # Ask user where to save the report
        from datetime import datetime
        default_filename = f"Raport_ButloDni_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        
        # Get default directory from APPROVED_FILE or use desktop
        try:
            from config import APPROVED_FILE
            default_dir = os.path.dirname(APPROVED_FILE)
        except:
            default_dir = os.path.expanduser("~/Desktop")
        
        default_path = os.path.join(default_dir, default_filename)
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Zapisz raport jako",
            default_path,
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if not output_path:
            # User cancelled
            return
        
        # Ensure .xlsx extension
        if not output_path.lower().endswith('.xlsx'):
            output_path += '.xlsx'
        
        # Start worker thread
        self.worker = ReportWorker(self.excel_handler, self.file_manager, filters, output_path)
        self.worker.finished.connect(self.on_report_finished)
        self.worker.start()
        
        # Show progress dialog
        self.progress_dialog = QProgressDialog("Generowanie raportu...", None, 0, 0, self)
        self.progress_dialog.setWindowTitle("Przetwarzanie")
        self.progress_dialog.show()
    
    def on_report_finished(self, success, message):
        """Handle report generation completion."""
        self.progress_dialog.close()
        
        if success:
            QMessageBox.information(self, "Sukces", message)
        else:
            QMessageBox.critical(self, "Błąd", message)