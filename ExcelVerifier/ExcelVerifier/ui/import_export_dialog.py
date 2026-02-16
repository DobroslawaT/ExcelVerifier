"""
Import/Export dialog for ExcelVerifier application.
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QMessageBox, QGroupBox, QRadioButton, QProgressDialog,
    QTextEdit, QCheckBox, QTabWidget, QWidget, QButtonGroup, QListView, 
    QTreeView, QListWidget, QDialogButtonBox, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
from datetime import datetime
import os

from core.import_export import ImportExportHandler


class ImportExportWorker(QThread):
    """Worker thread for import/export operations."""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)
    
    def __init__(self, operation, **kwargs):
        super().__init__()
        self.operation = operation
        self.kwargs = kwargs
        self.handler = ImportExportHandler()
    
    def run(self):
        try:
            if self.operation == "export":
                success, message = self.handler.export_all_data(self.kwargs['output_path'])
            elif self.operation == "import":
                success, message = self.handler.import_all_data(self.kwargs['zip_path'], self.kwargs.get('merge', False))
            elif self.operation == "import_excel":
                success, message = self.handler.import_from_excel_file(self.kwargs['excel_path'])
            else:
                success, message = False, "Nieznana operacja"
            
            self.finished.emit(success, message)
        except Exception as e:
            self.finished.emit(False, f"Błąd: {str(e)}")


class ImportExportDialog(QDialog):
    """Dialog for importing and exporting application data."""
    # Signal emitted when data is imported/exported to refresh main window
    data_refreshed = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Import / Export")
        self.resize(540, 420)
        self.worker = None
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(25)
        
        # Title
        title = QLabel("Zarządzanie Danymi")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #1F2937;")
        layout.addWidget(title)
        
        # === SYNCHRONIZATION SECTION ===
        sync_label = QLabel("🔄 Synchronizacja")
        sync_label.setStyleSheet("font-size: 13px; font-weight: 600; color: #6B7280; margin-bottom: 5px;")
        layout.addWidget(sync_label)
        
        sync_grid = QHBoxLayout()
        sync_grid.setSpacing(12)
        
        export_btn = self._create_action_card("📤", "Eksportuj", "#2563EB", self.export_data)
        import_zip_btn = self._create_action_card("📥", "Importuj", "#10B981", self.import_zip)
        
        sync_grid.addWidget(export_btn)
        sync_grid.addWidget(import_zip_btn)
        
        layout.addLayout(sync_grid)
        
        # === IMPORT SOURCES SECTION ===
        layout.addSpacing(10)
        
        sources_label = QLabel("📂 Importuj Źródła Danych")
        sources_label.setStyleSheet("font-size: 13px; font-weight: 600; color: #6B7280; margin-bottom: 5px;")
        layout.addWidget(sources_label)
        
        excel_grid = QHBoxLayout()
        excel_grid.setSpacing(0)
        excel_btn = self._create_action_card("📊", "Lista Zatwierdzonych (Excel)", "#8B5CF6", self.import_excel)
        excel_grid.addWidget(excel_btn)
        
        layout.addLayout(excel_grid)
        
        # === INFO TEXT ===
        layout.addSpacing(15)
        
        info = QLabel(
            "💡 Eksportuj: backup lub synchronizacja z inną aplikacją\n"
            "📊 Lista Zatwierdzonych (Excel): plik Excel ze ścieżkami → auto-kopiowanie plików + zdjęć"
        )
        info.setWordWrap(True)
        info.setStyleSheet("font-size: 11px; color: #9CA3AF; line-height: 1.5;")
        layout.addWidget(info)
        
        layout.addStretch()
        
        # === CLOSE BUTTON ===
        close_btn = QPushButton("Zamknij")
        close_btn.setFixedHeight(35)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: white;
                color: #6B7280;
                border: 1px solid #D1D5DB;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                font-weight: 500;
            }
            QPushButton:hover { 
                border-color: #9CA3AF;
                background-color: #F9FAFB;
            }
        """)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)
    
    def _create_action_card(self, icon, label, color, callback):
        """Create a simple action card button."""
        card = QPushButton()
        card.setMinimumHeight(85)
        card.setCursor(Qt.PointingHandCursor)
        card.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                border: none;
                border-radius: 8px;
                padding: 15px;
                color: white;
            }}
            QPushButton:hover {{
                background-color: {self._darken_color(color)};
            }}
            QPushButton:pressed {{
                background-color: {self._darken_color(color)};
            }}
        """)
        
        layout = QVBoxLayout()
        card.setLayout(layout)
        layout.setSpacing(5)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(10, 10, 10, 10)
        
        icon_label = QLabel(icon)
        icon_label.setStyleSheet("font-size: 28px; background: transparent; color: white;")
        icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(icon_label)
        
        text_label = QLabel(label)
        text_label.setStyleSheet("font-size: 13px; font-weight: 600; background: transparent; color: white; text-align: center;")
        text_label.setAlignment(Qt.AlignCenter)
        text_label.setWordWrap(True)
        layout.addWidget(text_label)
        
        card.clicked.connect(callback)
        
        # Make button expand horizontally
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        return card
    
    def _darken_color(self, hex_color, factor=0.1):
        """Darken a hex color for hover effect."""
        color_map = {
            "#2563EB": "#1D4ED8",
            "#10B981": "#059669",
            "#8B5CF6": "#7C3AED",
            "#F59E0B": "#D97706"
        }
        return color_map.get(hex_color, hex_color)
    
    def export_data(self):
        """Export all application data to ZIP."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"ExcelVerifier_Backup_{timestamp}.zip"
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Zapisz Archiwum Eksportu",
            default_name,
            "ZIP Files (*.zip)"
        )
        
        if not output_path:
            return
        
        # Show progress dialog
        progress = QProgressDialog("Eksportowanie danych...", None, 0, 0, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowTitle("Eksport")
        progress.show()
        
        # Run export in worker thread
        self.worker = ImportExportWorker("export", output_path=output_path)
        self.worker.finished.connect(lambda success, msg: self._on_operation_finished(success, msg, progress))
        self.worker.start()
    
    def import_zip(self):
        """Import data from ZIP archive."""
        zip_path, _ = QFileDialog.getOpenFileName(
            self,
            "Wybierz Archiwum ZIP",
            "",
            "ZIP Files (*.zip)"
        )
        
        if not zip_path:
            return
        
        # Ask about merge mode with custom buttons
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Tryb Importu")
        msg_box.setText("Jak chcesz zaimportować dane?")
        msg_box.setInformativeText(
            "Zastąp - usuń wszystkie obecne dane i zastąp nimi\n"
            "Połącz - dodaj do istniejących danych (pomija duplikaty)"
        )
        
        replace_btn = msg_box.addButton("Zastąp", QMessageBox.DestructiveRole)
        merge_btn = msg_box.addButton("Połącz", QMessageBox.AcceptRole)
        cancel_btn = msg_box.addButton("Anuluj", QMessageBox.RejectRole)
        
        msg_box.exec_()
        clicked = msg_box.clickedButton()
        
        if clicked == cancel_btn:
            return
        
        merge = (clicked == merge_btn)
        
        # Show progress dialog
        progress = QProgressDialog("Importowanie danych...", None, 0, 0, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowTitle("Import")
        progress.show()
        
        # Run import in worker thread
        self.worker = ImportExportWorker("import", zip_path=zip_path, merge=merge)
        self.worker.finished.connect(lambda success, msg: self._on_operation_finished(success, msg, progress))
        self.worker.start()
    
    def import_excel(self):
        """Import from ApprovedRecords.xlsx file - automatically copies Excel files and images."""
        excel_path, _ = QFileDialog.getOpenFileName(
            self,
            "Wybierz plik Excel ze ścieżkami do plików (np. ApprovedRecords.xlsx)",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if not excel_path:
            return
        
        # Show info about what will happen
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Potwierdź Import")
        msg_box.setText("System automatycznie:")
        msg_box.setInformativeText(
            "1. Przeczyta ścieżki plików Excel z wybranego pliku\n"
            "2. Znajdzie i skopiuje pliki Excel\n"
            "3. Znajdzie i skopiuje powiązane zdjęcia\n"
            "4. Doda wszystko do bazy danych\n\n"
            "Kontynuować?"
        )
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.Yes)
        
        if msg_box.exec_() == QMessageBox.No:
            return
        
        # Show progress dialog
        progress = QProgressDialog("Wczytywanie ścieżek i kopiowanie plików...", None, 0, 0, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowTitle("Import")
        progress.show()
        
        # Run import in worker thread
        self.worker = ImportExportWorker("import_excel", excel_path=excel_path)
        self.worker.finished.connect(lambda success, msg: self._on_operation_finished(success, msg, progress))
        self.worker.start()
    
    def _on_operation_finished(self, success, message, progress_dialog):
        """Handle operation completion."""
        progress_dialog.close()
        
        if success:
            QMessageBox.information(self, "Sukces", message)
            # Emit signal to refresh main window
            self.data_refreshed.emit()
            # Close the dialog after successful import
            self.accept()
        else:
            QMessageBox.critical(self, "Błąd", message)
