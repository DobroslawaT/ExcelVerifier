import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QPushButton, 
    QLabel, QTableWidget, QTableWidgetItem, QSizePolicy, QMessageBox, 
    QDialog, QLineEdit, QApplication, QProgressDialog
)
from PyQt5.QtGui import QPixmap, QFont, QColor, QBrush
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import shutil

# Import your new separate logic classes
from core.excel_handler import ExcelHandler
from core.file_manager import FileManager
from core.image_transformer import ImageTransformer
import config


class ReprocessWorker(QThread):
    """Worker thread to reprocess an image through AI."""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)  # success, message
    
    def __init__(self, image_path, output_folder):
        super().__init__()
        self.image_path = image_path
        self.output_folder = output_folder
    
    def run(self):
        try:
            self.progress.emit("Starting AI reprocessing...")
            transformer = ImageTransformer()
            output_path, highlighted = transformer.process_image_file(self.image_path, self.output_folder)
            msg = f"Reprocessed successfully: {os.path.basename(output_path)}"
            if highlighted > 0:
                msg += f" ({highlighted} discrepancies found)"
            self.finished.emit(True, msg)
        except Exception as e:
            self.finished.emit(False, f"Reprocessing failed: {str(e)}")


class VerifyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Verification App")

        # Initialize Logic Classes (The "Brain")
        self.excel_handler = ExcelHandler()
        self.file_manager = FileManager()
        
        # Internal state
        self.current_report_index = 0
        self.unapproved_reports = []
        self.current_image_path = None

        # Setup UI (The "Face")
        self.init_ui()
        
        # Load initial data
        self.load_unapproved_list()

    def init_ui(self):
        """Builds all the visual elements."""
        # Main Layout
        main_layout = QHBoxLayout()
        self.setLayout(main_layout)
        
        # Splitter (Draggable divider)
        self.splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(self.splitter)

        # --- LEFT SIDE: Table ---
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True) # Adds zebra striping
        self.table.setShowGrid(True)
        # Remove the default frame to let CSS handle borders
        self.table.setFrameShape(QTableWidget.NoFrame) 
        self.splitter.addWidget(self.table)

        # --- RIGHT SIDE: Panel ---
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        right_panel.setLayout(right_layout)
        self.splitter.addWidget(right_panel)

        # Image Label
        self.image_label = QLabel("No image loaded")
        self.image_label.setObjectName("ImageContainer") # <--- ID for CSS
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # Click event to show full image (using lambda to pass current path)
        self.image_label.mousePressEvent = lambda e: self.show_full_image()
        right_layout.addWidget(self.image_label)
        
        right_layout.addSpacing(12)

        # Action Buttons
        # Save -> Primary (Blue)
        self.save_btn = self._create_styled_button("Save Changes", 226)
        self.save_btn.setObjectName("PrimaryBtn")  
        
        # Approve -> Success (Green)
        self.approve_btn = self._create_styled_button("Approve", 220)
        self.approve_btn.setObjectName("SuccessBtn") 

        # Delete -> Danger (Red)
        self.delete_btn = self._create_styled_button("Delete", 220)
        self.delete_btn.setObjectName("DangerBtn")
        
        # Reprocess -> Info (Orange/Yellow)
        self.reprocess_btn = self._create_styled_button("Reprocess AI", 220)
        self.reprocess_btn.setObjectName("InfoBtn")

        # Refresh -> Standard (White/Gray)
        self.refresh_btn = self._create_styled_button("Refresh", 120)
        
        # Connect Signals
        self.save_btn.clicked.connect(self.save_changes)
        self.approve_btn.clicked.connect(self.approve_current_report)
        self.delete_btn.clicked.connect(self.delete_current_report)
        self.reprocess_btn.clicked.connect(self.reprocess_current_report)
        self.refresh_btn.clicked.connect(self.refresh)

        # Add buttons to layout
        right_layout.addWidget(self.save_btn, alignment=Qt.AlignHCenter)
        right_layout.addSpacing(6)
        right_layout.addWidget(self.approve_btn, alignment=Qt.AlignHCenter)
        right_layout.addSpacing(6)
        right_layout.addWidget(self.delete_btn, alignment=Qt.AlignHCenter)
        right_layout.addSpacing(6)
        right_layout.addWidget(self.reprocess_btn, alignment=Qt.AlignHCenter)
        right_layout.addSpacing(6)
        right_layout.addWidget(self.refresh_btn, alignment=Qt.AlignHCenter)
        
        right_layout.addStretch(1) # Push navigation to bottom

        # Navigation (Prev/Next)
        nav_widget = QWidget()
        nav_layout = QHBoxLayout()
        nav_widget.setLayout(nav_layout)
        
        self.prev_btn = QPushButton("◀")
        self.next_btn = QPushButton("▶")
        self.nav_label = QLabel("0/0")
        
        self.prev_btn.setFixedWidth(40)
        self.next_btn.setFixedWidth(40)
        self.prev_btn.clicked.connect(self.prev_report)
        self.next_btn.clicked.connect(self.next_report)

        nav_layout.addWidget(self.prev_btn)
        nav_layout.addWidget(self.nav_label, alignment=Qt.AlignCenter)
        nav_layout.addWidget(self.next_btn)
        
        right_layout.addWidget(nav_widget)

        # Zatwierdzone Button
        self.view_approved_btn = self._create_styled_button("Zatwierdzone", 140)
        self.view_approved_btn.clicked.connect(self.show_approved_dialog)
        right_layout.addWidget(self.view_approved_btn, alignment=Qt.AlignHCenter)

        # Initial Splitter Sizes (70% Table, 30% Panel)
        self.splitter.setStretchFactor(0, 7)
        self.splitter.setStretchFactor(1, 3)

    def _create_styled_button(self, text, width):
        """Helper to create consistent buttons."""
        btn = QPushButton(text)
        btn.setFixedWidth(width)
        btn.setFixedHeight(40)
        font = QFont()
        font.setPointSize(11)
        btn.setFont(font)
        return btn

    def load_unapproved_list(self):
        """Uses FileManager to find work."""
        self.unapproved_reports = self.file_manager.get_unapproved_reports()
        
        if not self.unapproved_reports:
            self.nav_label.setText("0/0")
            QMessageBox.information(self, "Zrobione", "Nie znaleziono niezatwierdzonych raportów!")
            # Optionally load a fallback or disable UI
            return

        self.current_report_index = 0
        self.load_current_report()

    def load_current_report(self):
        """Loads data AND formatting from ExcelHandler."""
        if not self.unapproved_reports:
            return

        path = self.unapproved_reports[self.current_report_index]
        self.nav_label.setText(f"{self.current_report_index + 1}/{len(self.unapproved_reports)}")

        try:
            # 1. Load Data
            df = self.excel_handler.load_file(path)
            
            # 2. Load Formatting (The new step)
            style_map = self.excel_handler.get_formatting()

        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie można załadować Excela:\n{e}")
            return

        self.table.blockSignals(True)
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        
        headers = [f"Col {i+1}" for i in range(len(df.columns))]
        self.table.setHorizontalHeaderLabels(headers)

        # 3. Fill cells and Apply Colors
        for r in range(len(df)):
            for c in range(len(df.columns)):
                # Set Text
                val = str(df.iloc[r, c]) if df.iloc[r, c] is not None else ""
                item = QTableWidgetItem(val)

                # Set Colors if they exist in our style_map
                if (r, c) in style_map:
                    style = style_map[(r, c)]
                    
                    # Background Color
                    if style['bg']:
                        item.setBackground(QBrush(QColor(style['bg'])))
                    
                    # Text Color
                    if style['fg']:
                        item.setForeground(QBrush(QColor(style['fg'])))

                self.table.setItem(r, c, item)
        
        self.table.blockSignals(False)

        # 4. Load Image (using Config default or derived path)
        self.current_image_path = config.DEFAULT_IMAGE # Or logic to find specific image
        if os.path.exists(self.current_image_path):
            pix = QPixmap(self.current_image_path)
            # Scale for preview (keep aspect ratio)
            scaled = pix.scaled(500, 500, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_label.setPixmap(scaled)
        else:
            self.image_label.setText("Image not found")

    def save_changes(self):
        """Scrapes the UI table and sends data to ExcelHandler to save."""
        # 1. Scrape data from QTableWidget
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        data_from_table = [] 
        
        for r in range(rows):
            row_data = []
            for c in range(cols):
                item = self.table.item(r, c)
                text = item.text() if item else ""
                row_data.append(text)
            data_from_table.append(row_data)
        
        # 2. Pass data to logic layer
        try:
            # The handler will update the excel file, run the red-highlight logic,
            # and update database
            self.excel_handler.save_data(data_from_table)
            
            # Visual feedback
            QMessageBox.information(self, "Zapisano", "Zmiany zapisane pomyślnie.\nLogika zweryfikowana.")
            
            # Optional: Reload to show any new red highlights the Handler added
            self.load_current_report() 
            
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", f"An error occurred:\n{str(e)}")

    def approve_current_report(self):
        """Calls logic to mark file as approved."""
        if not self.unapproved_reports:
            return

        current_path = self.unapproved_reports[self.current_report_index]
        
        try:
            # 1. Save any pending edits first
            self.save_changes() 
            
            # 2. Call handler to write to database
            filename = os.path.basename(current_path)
            print(f"\n=== APPROVE START ===")
            print(f"DEBUG: Current index: {self.current_report_index}")
            print(f"DEBUG: Approving file: {filename}")
            print(f"DEBUG: Full path: {current_path}")
            print(f"DEBUG: List before pop: {[os.path.basename(p) for p in self.unapproved_reports]}")
            
            # Simple parsing logic (can be moved to utils)
            name_no_ext = os.path.splitext(filename)[0]
            date_part = name_no_ext[:10]
            company_part = name_no_ext[11:]
            
            self.excel_handler.approve_report(filename, date_part, company_part, current_path)
            print(f"DEBUG: Approval saved to database")
            
            QMessageBox.information(self, "Zatwierdzono", "Raport został zatwierdzony.")
            
            # 3. Manually remove the current file from the list instead of re-querying
            # This is more reliable than checking if it was added to database
            self.unapproved_reports.pop(self.current_report_index)
            print(f"DEBUG: Removed file at index {self.current_report_index}")
            print(f"DEBUG: List after pop: {[os.path.basename(p) for p in self.unapproved_reports]}")
            print(f"DEBUG: Remaining count: {len(self.unapproved_reports)}")
            
            if not self.unapproved_reports:
                self.nav_label.setText("0/0")
                QMessageBox.information(self, "Gotowe", "Brak niezatwierdzonych raportów!")
                print(f"=== APPROVE END (No more reports) ===\n")
                return
            
            # Keep the same index (now points to next file, or cap to last if we were at end)
            if self.current_report_index >= len(self.unapproved_reports):
                print(f"DEBUG: Index {self.current_report_index} >= list length {len(self.unapproved_reports)}, capping to {len(self.unapproved_reports) - 1}")
                self.current_report_index = len(self.unapproved_reports) - 1
            
            self.current_report_index = max(0, self.current_report_index)
            print(f"DEBUG: Final index to load: {self.current_report_index}")
            print(f"DEBUG: File to load: {os.path.basename(self.unapproved_reports[self.current_report_index])}")
            print(f"=== APPROVE END (loading next) ===\n")
            
            self.load_current_report()
            
        except Exception as e:
            print(f"=== APPROVE ERROR ===\n{e}\n")
            QMessageBox.critical(self, "Błąd", f"Nie udało się zatwierdzić:\n{e}")

    def delete_current_report(self):
        """Delete the current report file (works for unapproved reports)."""
        if not self.unapproved_reports:
            return
        
        current_path = self.unapproved_reports[self.current_report_index]
        filename = os.path.basename(current_path)
        
        # Confirm deletion
        reply = QMessageBox.warning(
            self,
            "Usuń raport",
            f"Czy na pewno chcesz usunąć:\n{filename}?\n\nTej operacji nie można cofnąć.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                # Delete the file
                if os.path.exists(current_path):
                    os.remove(current_path)
                    QMessageBox.information(self, "Usunięto", f"Raport usunięty: {filename}")
                    # Refresh list
                    self.load_unapproved_list()
                else:
                    QMessageBox.warning(self, "Nie znaleziono", "Nie znaleziono pliku.")
            except Exception as e:
                QMessageBox.critical(self, "Błąd", f"Nie udało się usunąć raportu:\n{e}")
    
    def delete_approved_report(self, filename):
        """Delete an approved report from database."""
        # Confirm deletion
        reply = QMessageBox.warning(
            self,
            "Usuń zatwierdzony raport",
            f"Czy na pewno chcesz usunąć zatwierdzony rekord:\n{filename}?\n\nTej operacji nie można cofnąć.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                self.excel_handler.delete_approved_record(filename)
                QMessageBox.information(self, "Usunięto", f"Zatwierdzony rekord usunięty: {filename}")
            except Exception as e:
                QMessageBox.critical(self, "Błąd", f"Nie udało się usunąć zatwierdzonego rekordu:\n{e}")

    def reprocess_current_report(self):
        """Reprocess the current report through AI."""
        if not self.unapproved_reports:
            return
        
        current_path = self.unapproved_reports[self.current_report_index]
        filename = os.path.basename(current_path)
        
        # Find the original image (usually has same name but different extension)
        image_name = os.path.splitext(filename)[0]
        image_dir = os.path.dirname(current_path)
        
        image_path = None
        for ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']:
            potential_path = os.path.join(image_dir, image_name + ext)
            if os.path.exists(potential_path):
                image_path = potential_path
                break
        
        if not image_path:
            QMessageBox.warning(self, "Nie znaleziono", f"Nie znaleziono oryginalnego obrazu dla: {image_name}")
            return
        
        # Show progress dialog
        self.progress_dialog = QProgressDialog(
            "Przetwarzanie przez AI...",
            "Anuluj",
            0, 0,
            self
        )
        self.progress_dialog.setWindowModality(Qt.NonModal)
        self.progress_dialog.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.progress_dialog.show()
        
        # Create and start worker thread
        self.reprocess_worker = ReprocessWorker(image_path, config.TRANSFORM_DIRECTORY)
        self.reprocess_worker.finished.connect(self.on_reprocess_finished)
        self.reprocess_worker.start()
    
    def on_reprocess_finished(self, success, message):
        """Handle reprocessing completion."""
        self.progress_dialog.close()
        
        if success:
            QMessageBox.information(self, "Przetwarzanie zakończone", message)
            # Refresh the list to show updated files
            self.load_unapproved_list()
        else:
            QMessageBox.critical(self, "Przetwarzanie nie powiodło się", message)

    def refresh(self):
        # Reload the current file from disk (discarding unsaved UI changes)
        self.load_current_report()

    def prev_report(self):
        if self.current_report_index > 0:
            self.current_report_index -= 1
            self.load_current_report()

    def next_report(self):
        if self.current_report_index < len(self.unapproved_reports) - 1:
            self.current_report_index += 1
            self.load_current_report()

    def show_full_image(self):
        if self.current_image_path and os.path.exists(self.current_image_path):
            # You can keep your Dialog logic here or move it to ui/dialogs.py
            dlg = QDialog(self)
            dlg.setWindowTitle("Full Image")
            dlg.showMaximized()
            layout = QVBoxLayout(dlg)
            lbl = QLabel()
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setPixmap(QPixmap(self.current_image_path).scaled(
                dlg.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
            ))
            layout.addWidget(lbl)
            dlg.exec_()

    def show_approved_dialog(self):
        QMessageBox.information(self, "TODO", "Move your Approved Dialog logic to ui/dialogs.py and instantiate it here!")