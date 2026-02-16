import json
from pathlib import Path
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QLineEdit, QFileDialog, QGroupBox, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QFont

from ui.company_db_dialog import CompanyDbDialog


class SettingsDialog(QDialog):
    """Dialog for configuring application directories."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ustawienia - Konfiguracja Katalogów")
        self.setMinimumWidth(600)
        self.setModal(True)
        
        # Create icon from settings emoji
        pixmap = QPixmap(64, 64)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        font = QFont()
        font.setPointSize(40)
        painter.setFont(font)
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "\u2699")
        painter.end()
        self.setWindowIcon(QIcon(pixmap))
        
        # Get project root (3 levels up from this file)
        self.project_root = Path(__file__).parent.parent.parent.parent
        self.settings_file = self.project_root / "settings.json"
        self.settings = self.load_settings()
        
        self.init_ui()
        self.load_current_values()
    
    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Reports Directory (Unapproved)
        reports_group = QGroupBox("Katalog Raportów (Niezatwierdzone Pliki)")
        reports_layout = QVBoxLayout()
        reports_group.setLayout(reports_layout)
        
        reports_label = QLabel("Katalog, w którym przechowywane są niezatwierdzone raporty Excel:")
        reports_layout.addWidget(reports_label)
        
        reports_path_layout = QHBoxLayout()
        self.reports_path_input = QLineEdit()
        self.reports_path_input.setReadOnly(True)
        reports_path_layout.addWidget(self.reports_path_input)
        
        reports_browse_btn = QPushButton("Przeglądaj...")
        reports_browse_btn.clicked.connect(lambda: self.browse_directory("reports"))
        reports_path_layout.addWidget(reports_browse_btn)
        
        reports_layout.addLayout(reports_path_layout)
        layout.addWidget(reports_group)
        
        # Approved Directory
        approved_group = QGroupBox("Katalog Ztwierdzonych Raportów")
        approved_layout = QVBoxLayout()
        approved_group.setLayout(approved_layout)
        
        approved_label = QLabel("Katalog, w którym przechowywany jest plik Excel z zatwierdzonymi raportami:")
        approved_layout.addWidget(approved_label)
        
        approved_path_layout = QHBoxLayout()
        self.approved_path_input = QLineEdit()
        self.approved_path_input.setReadOnly(True)
        approved_path_layout.addWidget(self.approved_path_input)
        
        approved_browse_btn = QPushButton("Przeglądaj...")
        approved_browse_btn.clicked.connect(lambda: self.browse_directory("approved"))
        approved_path_layout.addWidget(approved_browse_btn)
        
        approved_layout.addLayout(approved_path_layout)
        layout.addWidget(approved_group)
        
        # Transform Output Directory
        transform_group = QGroupBox("Wyjście Zdjęcie na Excel")
        transform_layout = QVBoxLayout()
        transform_group.setLayout(transform_layout)
        
        transform_label = QLabel("Katalog, w którym zostaną utworzone przekształcone pliki Excel:")
        transform_layout.addWidget(transform_label)
        
        transform_path_layout = QHBoxLayout()
        self.transform_path_input = QLineEdit()
        self.transform_path_input.setReadOnly(True)
        transform_path_layout.addWidget(self.transform_path_input)
        
        transform_browse_btn = QPushButton("Przeglądaj...")
        transform_browse_btn.clicked.connect(lambda: self.browse_directory("transform"))
        transform_path_layout.addWidget(transform_browse_btn)
        
        transform_layout.addLayout(transform_path_layout)
        layout.addWidget(transform_group)

        # Company Database
        company_group = QGroupBox("Baza firm do weryfikacji")
        company_layout = QVBoxLayout()
        company_group.setLayout(company_layout)

        company_label = QLabel("Dodaj lub usuń firmy (nazwa + NIP) używane do weryfikacji:")
        company_layout.addWidget(company_label)

        company_btn_row = QHBoxLayout()
        company_btn = QPushButton("Dodaj / usuń firmę")
        company_btn.clicked.connect(self.open_company_db)
        company_btn_row.addWidget(company_btn)
        company_btn_row.addStretch(1)
        company_layout.addLayout(company_btn_row)

        layout.addWidget(company_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        save_btn = QPushButton("Zapisz")
        save_btn.clicked.connect(self.save_and_close)
        save_btn.setMinimumWidth(100)
        button_layout.addWidget(save_btn)
        
        cancel_btn = QPushButton("Anuluj")
        cancel_btn.clicked.connect(self.reject)
        cancel_btn.setMinimumWidth(100)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
    
    def load_settings(self):
        """Load settings from JSON file."""
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        
        # Default settings (relative paths)
        return {
            "reports_directory": "Reports/Niezatwierdzone",
            "approved_directory": "Reports/Zatwierdzone",
            "transform_directory": "Reports"
        }
    
    def save_settings(self):
        """Save settings to JSON file."""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=4)
            return True
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie udało się zapisać ustawień: {str(e)}")
            return False
    
    def resolve_display_path(self, path_str):
        """Resolve path for display (show absolute path in UI)."""
        path = Path(path_str)
        if path.is_absolute():
            return str(path)
        else:
            return str((self.project_root / path).resolve())
    
    def load_current_values(self):
        """Load current settings into input fields."""
        reports = self.settings.get("reports_directory", "")
        approved = self.settings.get("approved_directory", "")
        transform = self.settings.get("transform_directory", "")
        
        self.reports_path_input.setText(self.resolve_display_path(reports))
        self.approved_path_input.setText(self.resolve_display_path(approved))
        self.transform_path_input.setText(self.resolve_display_path(transform))
    
    def make_relative_if_possible(self, abs_path):
        """Convert absolute path to relative if it's under project root."""
        try:
            abs_path = Path(abs_path)
            rel_path = abs_path.relative_to(self.project_root)
            return str(rel_path).replace('\\', '/')
        except ValueError:
            # Path is outside project root, keep as absolute
            return str(abs_path).replace('\\', '/')
    
    def browse_directory(self, dir_type):
        """Open directory browser."""
        if dir_type == "reports":
            current = self.reports_path_input.text()
            title = "Select Reports Directory (Unapproved)"
        elif dir_type == "approved":
            current = self.approved_path_input.text()
            title = "Select Approved Reports Directory"
        else:  # transform
            current = self.transform_path_input.text()
            title = "Select Transform Output Directory"
        
        directory = QFileDialog.getExistingDirectory(
            self,
            title,
            current if current else str(self.project_root)
        )
        
        if directory:
            # Store as relative path if possible
            stored_path = self.make_relative_if_possible(directory)
            
            if dir_type == "reports":
                self.reports_path_input.setText(directory)
                self.settings["reports_directory"] = stored_path
            elif dir_type == "approved":
                self.approved_path_input.setText(directory)
                self.settings["approved_directory"] = stored_path
            else:  # transform
                self.transform_path_input.setText(directory)
                self.settings["transform_directory"] = stored_path
    
    def save_and_close(self):
        """Save settings and close dialog."""
        if self.save_settings():
            QMessageBox.information(
                self, 
                "Settings Saved", 
                "Settings have been saved successfully.\n\nRestart the application for changes to take effect."
            )
            self.accept()

    def open_company_db(self):
        dlg = CompanyDbDialog(self)
        dlg.exec_()
    
    def get_settings(self):
        """Return current settings."""
        return self.settings
