import os
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, 
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, 
    QSizePolicy, QMessageBox, QTabWidget, QDialog, QHeaderView,
    QAbstractItemView, QToolButton, QApplication, QProgressDialog,
    QAbstractScrollArea, QComboBox, QCompleter, QStyledItemDelegate,
    QLineEdit
)
from PyQt5.QtGui import QPixmap, QFont, QColor, QBrush
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal

# Import logic classes
from ui.styles import STYLESHEET
from core.excel_handler import ExcelHandler
from core.file_manager import FileManager
from core.company_db import load_company_db
import config
from ui.dialogs import ApprovedReportsDialog, UnapprovedReportsDialog
from ui.TransformPicToExcelPage import TransformPage
from ui.GenerateReportPage import GenerateReportPage
from ui.settings_dialog import SettingsDialog
from ui.import_export_dialog import ImportExportDialog

# --- WORKER THREAD: Reprocess Single Image ---
class ReprocessWorker(QThread):
    finished = pyqtSignal(bool, str, str, int)  # success, message, output_path, highlighted
    
    def __init__(self, image_path, output_folder):
        super().__init__()
        self.image_path = image_path
        self.output_folder = output_folder
    
    def run(self):
        try:
            from core.image_transformer import ImageTransformer
            transformer = ImageTransformer()
            output_path, highlighted = transformer.process_image_file(self.image_path, self.output_folder)
            msg = f"Przetworzono pomyślnie: {os.path.basename(output_path)}"
            if highlighted > 0:
                msg += f"\n\n({highlighted} wykrytych rozbieżności)"
            self.finished.emit(True, msg, output_path, highlighted)
        except Exception as e:
            self.finished.emit(False, f"Przetwarzanie nie powiodło się: {str(e)}", "", 0)

# --- CUSTOM WIDGET: Auto-Resizing Image Label ---
class AutoResizingLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setAlignment(Qt.AlignCenter)
        self._pixmap = None

    def setPixmap(self, pixmap):
        self._pixmap = pixmap
        self.update_view()

    def resizeEvent(self, event):
        self.update_view()
        super().resizeEvent(event)

    def update_view(self):
        if self._pixmap and not self._pixmap.isNull():
            scaled = self._pixmap.scaled(
                self.size(), 
                Qt.KeepAspectRatio, 
                Qt.SmoothTransformation
            )
            super().setPixmap(scaled)
        else:
            super().setPixmap(QPixmap())


class OdbiorcaComboDelegate(QStyledItemDelegate):
    def __init__(self, company_list, parent=None):
        super().__init__(parent)
        self.company_list = company_list

    def createEditor(self, parent, option, index):
        if index.row() != 0 or index.column() != 1:
            return super().createEditor(parent, option, index)

        combo = QComboBox(parent)
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.addItems(self.company_list)
        combo.setMinimumWidth(420)
        combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        combo.view().setMinimumWidth(600)
        combo.view().setTextElideMode(Qt.ElideNone)

        completer = QCompleter(self.company_list, combo)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        combo.setCompleter(completer)

        line_edit = combo.lineEdit()
        if line_edit:
            line_edit.setPlaceholderText("Wpisz, aby wyszukac firmę...")
            line_edit.setClearButtonEnabled(True)

        return combo

    def setEditorData(self, editor, index):
        if isinstance(editor, QComboBox):
            editor.setCurrentText(index.data() or "")
        else:
            super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentText())
        else:
            super().setModelData(editor, model, index)

# --- 1. The Verification App Logic ---
class VerificationPage(QWidget):
    def __init__(self):
        super().__init__()
        
        # Initialize Logic Classes
        self.excel_handler = ExcelHandler()
        self.file_manager = FileManager()
        
        self.current_report_index = 0
        self.unapproved_reports = []
        self.current_image_path = None
        self.current_excel_path = None
        self.is_review_mode = False 
        self.approved_reports = [] 
        self.current_approved_index = 0 
        # Track if user manually adjusted the splitter; if so, don't override
        self.user_split_override = False
        # Preserve full header text to elide on resize
        self._full_file_name_text = "Excel: "
        self.company_list = []

        self.init_ui()
        self.load_unapproved_list()

    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)
        layout.setContentsMargins(0, 20, 20, 10)
        layout.setSpacing(0)
        
        # --- HEADER ---
        self.info_widget = QWidget()
        self.info_widget.setFixedHeight(32)
        info_layout = QHBoxLayout(self.info_widget)
        info_layout.setContentsMargins(12, 0, 12, 0)
        info_layout.setSpacing(5)
        
        self.file_icon_label = QLabel("📄")
        self.file_icon_label.setStyleSheet("font-size: 12px; color: #374151;")
        info_layout.addWidget(self.file_icon_label)
        
        self.file_name_label = QLabel("Nie załadowano pliku")
        self.file_name_label.setStyleSheet("""
            font-size: 15px; font-weight: 700; color: #2563EB; text-decoration: underline;
        """)
        # Allow shrinking instead of forcing window to grow; no wrap
        self.file_name_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.file_name_label.setWordWrap(False)
        self.file_name_label.setMinimumWidth(220)
        self.file_name_label.setCursor(Qt.PointingHandCursor)
        self.file_name_label.mousePressEvent = lambda e: self.open_excel_file()
        info_layout.addWidget(self.file_name_label)
        
        info_layout.addStretch(8)
        
        self.picture_name_label = QLabel("Brak zdjęcia")
        self.picture_name_label.setStyleSheet("font-size: 12px; color: #6B7280;")
        self.picture_name_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        info_layout.addWidget(self.picture_name_label)
        info_layout.addStretch(2)
        
        self.info_widget.setStyleSheet("""
            QWidget { background-color: #FFFFFF; border-bottom: 1px solid #E5E7EB; }
        """)
        layout.addWidget(self.info_widget)
        
        # --- EMPTY STATE ---
        self.empty_state = QWidget()
        empty_layout = QVBoxLayout(self.empty_state)
        empty_layout.setAlignment(Qt.AlignCenter)
        
        empty_icon = QLabel("📋")
        empty_icon.setStyleSheet("font-size: 64px;")
        empty_icon.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_icon)
        
        empty_title = QLabel("Brak Niezatwierdzonych Raportów")
        empty_title.setStyleSheet("font-size: 24px; font-weight: bold; color: #111827; margin-top: 20px;")
        empty_title.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_title)
        
        empty_desc = QLabel("Wszystkie raporty zostały zweryfikowane i zatwierdzone.\nSprawdź archiwum zatwierdzonych raportów.")
        empty_desc.setStyleSheet("font-size: 14px; color: #6B7280; margin-top: 10px; margin-bottom: 30px;")
        empty_desc.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_desc)
        
        view_approved_empty_btn = QPushButton("📁 Zatwierdzone Raporty")
        view_approved_empty_btn.setFixedSize(220, 50)
        view_approved_empty_btn.setCursor(Qt.PointingHandCursor)
        view_approved_empty_btn.setStyleSheet("""
            QPushButton { background-color: #2563EB; color: white; border: none; border-radius: 8px; font-size: 15px; font-weight: bold; }
            QPushButton:hover { background-color: #1D4ED8; }
        """)
        view_approved_empty_btn.clicked.connect(self.show_approved_dialog)
        empty_layout.addWidget(view_approved_empty_btn, alignment=Qt.AlignCenter)
        
        self.empty_state.hide()
        layout.addWidget(self.empty_state)
        
        # --- SPLITTER (Main Content) ---
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setHandleWidth(6)  # Make the splitter handle more visible/grabbable
        self.splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #D1D5DB;
            }
            QSplitter::handle:hover {
                background-color: #9CA3AF;
            }
        """)
        layout.addWidget(self.splitter)

        # LEFT SIDE: Table
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)
        left_panel.setLayout(left_layout)

        self.table = QTableWidget()
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)
        # Prevent auto-expansion from content: keep widths logic-controlled
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.setWordWrap(True)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        # Prevent contents from growing the window; allow scrolling instead
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setMinimumWidth(320)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustIgnored)
        self.table.setEditTriggers(
            QAbstractItemView.DoubleClicked |
            QAbstractItemView.SelectedClicked |
            QAbstractItemView.EditKeyPressed
        )
        self.table.itemChanged.connect(self._on_table_item_changed)
        left_layout.addWidget(self.table, 1)
        self.splitter.addWidget(left_panel)

        # RIGHT SIDE: Panel
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(10, 0, 0, 0)
        right_panel.setLayout(right_layout)
        right_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.splitter.addWidget(right_panel)
        
        # Set stretch factors: Table 3 parts (60%), Image 2 parts (40%)
        # 3:2 ratio = 60%:40% = Table:Image
        self.splitter.setStretchFactor(0, 3)
        self.splitter.setStretchFactor(1, 2)
        # Detect user manual changes
        self.splitter.splitterMoved.connect(self.on_splitter_moved)

        # Image
        self.image_label = AutoResizingLabel("Nie załadowano zdjęcia")
        self.image_label.setStyleSheet("border: 1px solid #E5E7EB; background-color: #F9FAFB;")
        self.image_label.setMinimumHeight(150)  # Ensure image doesn't shrink too small
        right_layout.addWidget(self.image_label, 1)  # Stretch factor 1

        # Controls
        controls_container = QWidget()
        controls_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)  # Always show controls
        controls_layout = QVBoxLayout(controls_container)
        controls_layout.setSpacing(10)
        controls_layout.setContentsMargins(0, 10, 0, 0)

        self.save_btn = self._create_styled_button("Zapisz Zmiany", 140)
        self.approve_btn = self._create_styled_button("Zatwierdź", 140)
        self.delete_btn = self._create_styled_button("Usuń", 140)
        self.reprocess_btn = self._create_styled_button("Przetwarzaj AI", 140)
        self.refresh_btn = self._create_styled_button("Odśwież", 140)

        self.save_btn.setStyleSheet("""
            QPushButton { background-color: #2563EB; color: white; border: none; border-radius: 8px; }
            QPushButton:hover { background-color: #1D4ED8; }
        """)
        self.approve_btn.setStyleSheet("""
            QPushButton { background-color: #10B981; color: white; border: none; border-radius: 8px; }
            QPushButton:hover { background-color: #059669; }
        """)
        self.delete_btn.setStyleSheet("""
            QPushButton { background-color: #EF4444; color: white; border: none; border-radius: 8px; }
            QPushButton:hover { background-color: #DC2626; }
        """)
        self.reprocess_btn.setStyleSheet("""
            QPushButton { background-color: #F59E0B; color: white; border: none; border-radius: 8px; }
            QPushButton:hover { background-color: #D97706; }
        """)
        self.refresh_btn.setStyleSheet("""
            QPushButton { background-color: transparent; color: #374151; border: 1px solid #D1D5DB; border-radius: 8px; }
            QPushButton:hover { background-color: #F9FAFB; }
        """)

        self.save_btn.clicked.connect(self.save_changes)
        self.approve_btn.clicked.connect(self.approve_current_report)
        self.delete_btn.clicked.connect(self.delete_current_report)
        self.reprocess_btn.clicked.connect(self.reprocess_current_report)
        self.refresh_btn.clicked.connect(self.refresh)

        actions_row = QHBoxLayout()
        actions_row.setSpacing(8)
        actions_row.addWidget(self.save_btn)
        actions_row.addWidget(self.approve_btn)
        actions_row.addWidget(self.delete_btn)
        actions_row.addWidget(self.reprocess_btn)
        
        controls_layout.addLayout(actions_row)
        controls_layout.addWidget(self.refresh_btn, alignment=Qt.AlignHCenter)

        nav_widget = QWidget()
        nav_layout = QHBoxLayout()
        nav_layout.setSpacing(6)
        nav_widget.setLayout(nav_layout)
        self.prev_btn = QPushButton("◀")
        self.next_btn = QPushButton("▶")
        self.nav_label = QLabel("0/0")
        self.prev_btn.setFixedWidth(40)
        self.next_btn.setFixedWidth(40)
        self.prev_btn.clicked.connect(self.prev_report)
        self.next_btn.clicked.connect(self.next_report)
        
        nav_layout.addStretch(1)
        nav_layout.addWidget(self.prev_btn)
        nav_layout.addWidget(self.nav_label)
        nav_layout.addWidget(self.next_btn)
        nav_layout.addStretch(1)
        controls_layout.addWidget(nav_widget)

        approved_nav_layout = QHBoxLayout()
        self.view_approved_btn = self._create_styled_button("Zatwierdzone", 165)
        self.view_approved_btn.clicked.connect(self.show_approved_dialog)
        
        self.select_unapproved_btn = self._create_styled_button("Niezatwierdzone", 180)
        self.select_unapproved_btn.clicked.connect(self.show_unapproved_dialog)
        
        self.back_to_unapproved_btn = self._create_styled_button("Niezatwierdzone", 190)
        self.back_to_unapproved_btn.clicked.connect(self.back_to_unapproved)
        self.back_to_unapproved_btn.hide()

        approved_nav_layout.addWidget(self.view_approved_btn)
        approved_nav_layout.addWidget(self.select_unapproved_btn)
        approved_nav_layout.addWidget(self.back_to_unapproved_btn)
        
        controls_layout.addLayout(approved_nav_layout)
        right_layout.addWidget(controls_container, 0)

        # --- SMART SPLITTER CONFIGURATION ---
        self.splitter.setCollapsible(0, False)
        self.splitter.setCollapsible(1, False)
        
        # Balanced stretch factors to allow free resizing
        self.splitter.setStretchFactor(0, 3)  # Table gets 60%
        self.splitter.setStretchFactor(1, 2)  # Panel gets 40%
        
        # Set minimum sizes to prevent collapsing but allow full range
        # Table min is already set to 320 above
        right_panel.setMinimumWidth(200)
        right_panel.setMaximumWidth(16777215)  # Remove any implicit max width

        # Clamp table width to its pane to avoid window growth
        QTimer.singleShot(0, self.clamp_table_width)

    def _create_styled_button(self, text, width):
        btn = QPushButton(text)
        btn.setFixedWidth(width)
        btn.setFixedHeight(40)
        font = QFont()
        font.setPointSize(11)
        btn.setFont(font)
        return btn
    
    def find_linked_image(self, excel_path):
        excel_dir = os.path.dirname(excel_path)
        excel_base = os.path.splitext(os.path.basename(excel_path))[0]
        
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']
        for ext in image_extensions:
            potential_image = os.path.join(excel_dir, excel_base + ext)
            if os.path.exists(potential_image):
                return potential_image
        
        if config.DEFAULT_IMAGE and os.path.exists(config.DEFAULT_IMAGE):
            return config.DEFAULT_IMAGE
        return None

    def load_unapproved_list(self):
        self.unapproved_reports = self.file_manager.get_unapproved_reports()
        if not self.unapproved_reports:
            self.nav_label.setText("0/0")
            self.show_empty_state()
            return
        self.show_normal_view()
        self.current_report_index = 0
        self.load_current_report()
    
    def show_empty_state(self):
        self.info_widget.hide()
        self.splitter.hide()
        self.empty_state.show()
        self.file_name_label.clear()
        self.picture_name_label.clear()
    
    def show_normal_view(self):
        self.empty_state.hide()
        self.splitter.show()
        self.info_widget.show()

    def on_splitter_moved(self, pos, index):
        # User adjusted the split; stop auto-enforcing default ratio
        self.user_split_override = True
        # Keep table width clamped to left pane
        self.clamp_table_width()

    def apply_default_split(self):
        # Enforce 60:40 (Table:Image) only if user hasn't changed it
        if self.user_split_override:
            return
        w = self.splitter.width()
        if w > 0:
            self.splitter.setSizes([int(w * 0.6), int(w * 0.4)])
            # After sizes are applied, ensure table width stays within left pane
            self.clamp_table_width()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Keep default ratio on first load/resize until user changes it
        self.apply_default_split()
        # Update header label elided text to avoid width growth
        self.update_file_name_label_display()
        # Always clamp table width to current left pane
        self.clamp_table_width()

    def clamp_table_width(self):
        """Ensure table width doesn't have restrictive maximum set."""
        try:
            # Don't restrict the table's maximum width - let the splitter handle sizing
            self.table.setMaximumWidth(16777215)  # Qt's maximum widget size
        except Exception:
            pass

    def update_file_name_label_display(self):
        try:
            fm = self.file_name_label.fontMetrics()
            # Use current label width or a reasonable fallback
            avail = max(self.file_name_label.width() - 10, 200)
            text = self._full_file_name_text or ""
            self.file_name_label.setText(fm.elidedText(text, Qt.ElideRight, avail))
        except Exception:
            pass

    def _parse_date_to_standard_format(self, date_value, fallback):
        """Parse date from Excel cell to yyyy-MM-dd format."""
        from datetime import datetime
        
        if not date_value:
            return fallback
        
        # If it's already a datetime object
        if hasattr(date_value, 'strftime'):
            return date_value.strftime('%Y-%m-%d')
        
        # If it's a string, try to parse various formats
        date_str = str(date_value).strip()
        
        # Try common formats
        formats = [
            '%Y-%m-%d',      # 2026-02-12
            '%d.%m.%Y',      # 12.02.2026 (Polish)
            '%d/%m/%Y',      # 12/02/2026
            '%d-%m-%Y',      # 12-02-2026
            '%Y.%m.%d',      # 2026.02.12
            '%d.%m.%y',      # 12.02.26
            '%d/%m/%y',      # 12/02/26
        ]
        
        for fmt in formats:
            try:
                parsed = datetime.strptime(date_str, fmt)
                return parsed.strftime('%Y-%m-%d')
            except ValueError:
                continue
        
        # If nothing worked, return fallback
        return fallback
    
    def load_current_report(self):
        if not self.unapproved_reports: return
        path = self.unapproved_reports[self.current_report_index]
        self.current_excel_path = path
        self.nav_label.setText(f"{self.current_report_index + 1}/{len(self.unapproved_reports)}")
        
        file_name = os.path.basename(path)
        self.file_icon_label.setText("📄")
        self._full_file_name_text = f"Excel: {file_name}"
        QTimer.singleShot(0, self.update_file_name_label_display)
        
        linked_image = self.find_linked_image(path)
        if linked_image:
            pic_name = os.path.basename(linked_image)
            self.picture_name_label.setText(f"🖼️ Zdjęcie: {pic_name}")
        else:
            self.picture_name_label.setText("🖼️ Brak zdjęcia")
        
        try:
            df = self.excel_handler.load_file(path)
            style_map = self.excel_handler.get_formatting()
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Załadowanie nie powiodło się: {e}")
            return

        self.table.blockSignals(True)
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([f"Col {i+1}" for i in range(len(df.columns))])

        for r in range(len(df)):
            for c in range(len(df.columns)):
                cell_value = df.iloc[r, c]
                # Handle NaN, None, and empty values
                if pd.isna(cell_value) or cell_value is None:
                    val = ""
                else:
                    val = str(cell_value)
                item = QTableWidgetItem(val)
                
                # Enable text wrapping for second column (index 1)
                if c == 1:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignTop)
                    item.setFlags(item.flags() | Qt.TextWordWrap)
                
                if (r, c) in style_map:
                    style = style_map[(r, c)]
                    if style['bg']: item.setBackground(QBrush(QColor(style['bg'])))
                    if style['fg']: item.setForeground(QBrush(QColor(style['fg'])))
                    if style.get('bold') or style.get('italic') or style.get('underline'):
                        font = item.font()
                        if style.get('bold'): font.setBold(True)
                        if style.get('italic'): font.setItalic(True)
                        if style.get('underline'): font.setUnderline(True)
                        item.setFont(font)
                self.table.setItem(r, c, item)

            self._apply_company_selector()
        
        # 1. Resize columns to fit content (initial), then cap widths
        self.table.resizeColumnsToContents()

        # 2. Limit second column (index 1 - Odbiorca) width to 420px for text wrapping
        if self.table.columnCount() > 1:
            self.table.setColumnWidth(1, 420)

        # 3. Cap other columns and add small padding without inflating window
        for c in range(self.table.columnCount()):
            if c == 1:
                continue
            current_width = self.table.columnWidth(c)
            # small padding, with max cap to avoid window growth
            new_width = min(current_width + 8, 220)
            self.table.setColumnWidth(c, new_width)
        
        # 4. Make first row (header) taller for better readability
        if self.table.rowCount() > 0:
            self.table.setRowHeight(0, 75)
        
        # 5. Auto-resize all rows to fit wrapped text content
        self.table.resizeRowsToContents()
        
        self.table.blockSignals(False)
        # Apply default 60:40 split after render unless user adjusted manually
        QTimer.singleShot(0, self.apply_default_split)
        QTimer.singleShot(0, self.clamp_table_width)

        # Load linked image
        self.current_image_path = self.find_linked_image(path)
        if self.current_image_path and os.path.exists(self.current_image_path):
            pix = QPixmap(self.current_image_path)
            self.image_label.setPixmap(pix)
        else:
            self.image_label.setText("No image available")
            self.image_label.setPixmap(None)

    def save_changes(self):
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        data = [] 
        for r in range(rows):
            row_data = []
            for c in range(cols):
                item = self.table.item(r, c)
                text = item.text() if item else ""
                # Ensure "nan" string is converted to empty
                if text.lower() == "nan":
                    text = ""
                row_data.append(text)
            data.append(row_data)
        
        try:
            self.excel_handler.save_data(data)
            
            # Check if file path changed (due to company name reorganization)
            new_file_path = self.excel_handler.file_path
            if new_file_path != self.current_excel_path:
                print(f"File reorganized: {self.current_excel_path} → {new_file_path}")
                self.current_excel_path = new_file_path
                # Update unapproved_reports list with new path
                if not self.is_review_mode and self.current_report_index < len(self.unapproved_reports):
                    self.unapproved_reports[self.current_report_index] = new_file_path
            
            if self.is_review_mode:
                if self.approved_reports and self.current_approved_index < len(self.approved_reports):
                    self.load_approved_report(self.approved_reports[self.current_approved_index])
            else:
                self.load_current_report()
            QMessageBox.information(self, "Zapisano", "Zmiany zapisane.")
        except Exception as e:
            QMessageBox.critical(self, "Błąd", str(e))

    def _apply_company_selector(self):
        if self.table.rowCount() == 0 or self.table.columnCount() < 2:
            return

        self.company_list = self._load_company_names()
        odbiorca_item = self.table.item(0, 1)
        odbiorca_value = odbiorca_item.text() if odbiorca_item else ""

        delegate = OdbiorcaComboDelegate(self.company_list, self.table)
        self.table.setItemDelegateForColumn(1, delegate)

        if odbiorca_item is None:
            odbiorca_item = QTableWidgetItem(odbiorca_value)
            self.table.setItem(0, 1, odbiorca_item)

        self._apply_odbiorca_validation(odbiorca_value)

    def _apply_odbiorca_validation(self, value):
        if not self.company_list:
            return

        normalized = self._normalize_company_name(value)
        matches = any(
            self._normalize_company_name(name) == normalized
            for name in self.company_list
        )

        item = self.table.item(0, 1)
        if item:
            if matches:
                item.setBackground(QBrush())
            else:
                item.setBackground(QBrush(QColor("#FEE2E2")))

    def _load_company_names(self):
        companies = load_company_db(config.COMPANY_DB_FILE)
        names = [item.get("name", "") for item in companies if item.get("name")]
        return sorted(set(names), key=lambda value: value.lower())

    def _normalize_company_name(self, value):
        return " ".join(str(value).strip().lower().split())

    def _on_table_item_changed(self, item):
        if item.row() == 0 and item.column() == 1:
            self._apply_odbiorca_validation(item.text())


    def approve_current_report(self):
        if not self.unapproved_reports: return
        try:
            self.save_changes()
            path = self.unapproved_reports[self.current_report_index]
            fname = os.path.basename(path)
            
            # Read actual date and company from Excel cells instead of filename
            ws = self.excel_handler.current_workbook.active
            date_from_excel = ws['D1'].value  # Data wystawienia
            company_from_excel = ws['B1'].value  # Odbiorca
            
            # Convert date to string format yyyy-MM-dd (like 2026-02-12)
            date_str = self._parse_date_to_standard_format(date_from_excel, fname[:10])
            company_str = str(company_from_excel).strip() if company_from_excel else fname[11:]
            
            self.excel_handler.approve_report(fname, date_str, company_str, path)
            QMessageBox.information(self, "Zatwierdzono", "Gotowe.")
            
            # Reload the approved list to make newly approved file available
            self.reload_approved_list()
            
            # Instead of reloading, just remove the approved file and load the next one
            self.unapproved_reports.pop(self.current_report_index)
            
            if not self.unapproved_reports:
                self.show_empty_state()
                QMessageBox.information(self, "Koniec", "Brak więcej niezatwierdzonych raportów!")
                return
            
            # Cap index if we were at the end
            if self.current_report_index >= len(self.unapproved_reports):
                self.current_report_index = len(self.unapproved_reports) - 1
            
            self.load_current_report()
        except Exception as e:
            QMessageBox.critical(self, "Błąd", str(e))

    def delete_current_report(self):
        """Delete the current report file and its associated image (works for both approved and unapproved)."""
        # Check if we're in review mode (viewing approved reports)
        if self.is_review_mode:
            if not self.approved_reports:
                return
            path = self.approved_reports[self.current_approved_index]
            is_approved = True
        else:
            if not self.unapproved_reports:
                return
            path = self.unapproved_reports[self.current_report_index]
            is_approved = False
        
        filename = os.path.basename(path)
        
        # Find linked image
        linked_image = self.find_linked_image(path)
        
        # Build confirmation message
        msg = f"Czy chcesz usunąć:\n{filename}"
        if is_approved:
            msg += "\n\n⚠️ To jest ZATWIERDZONY raport!"
            msg += "\nZostanie usunięty z bazy danych"
        if linked_image and os.path.exists(linked_image):
            msg += f"\n\nI powiązane zdjęcie:\n{os.path.basename(linked_image)}"
        msg += "\n\nTej operacji nie będzie można cofnąć."
        
        # Confirm deletion
        reply = QMessageBox.warning(
            self,
            "Usuń Raport",
            msg,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            deleted_items = []
            errors = []
            
            try:
                # Delete Excel file
                if os.path.exists(path):
                    os.remove(path)
                    deleted_items.append(filename)
                else:
                    errors.append(f"Plik Excel nie znaleziony: {filename}")
                
                # Delete linked image
                if linked_image and os.path.exists(linked_image):
                    os.remove(linked_image)
                    deleted_items.append(os.path.basename(linked_image))
                
                # If approved report, also delete from database
                if is_approved:
                    try:
                        from core.database_handler import DatabaseHandler
                        from config import DATABASE_FILE
                        db = DatabaseHandler(DATABASE_FILE)
                        db.delete_reporting_data_by_filename(filename)
                        db.delete_approved_record(filename)
                        deleted_items.append("Wpis w bazie danych")
                    except Exception as e:
                        errors.append(f"Nie można usunąć z bazy danych: {e}")
                
                # Show result
                if deleted_items:
                    msg = "Usunięte:\n" + "\n".join(f"  • {item}" for item in deleted_items)
                    if errors:
                        msg += "\n\n⚠️ Ostrzeżenia:\n" + "\n".join(f"  • {err}" for err in errors)
                    QMessageBox.information(self, "Usunięto", msg)
                    
                    # Refresh appropriate list
                    if is_approved:
                        # Reload the approved list from Excel to reflect deletion
                        self.reload_approved_list()
                        
                        # Remove from in-memory approved list and load next or go back to unapproved
                        if path in self.approved_reports:
                            self.approved_reports.remove(path)
                        
                        if self.approved_reports:
                            # Adjust index if needed
                            if self.current_approved_index >= len(self.approved_reports):
                                self.current_approved_index = max(0, len(self.approved_reports) - 1)
                            # Load next report
                            self.load_approved_report(self.approved_reports[self.current_approved_index])
                        else:
                            # No more approved reports, go back to unapproved view
                            self.back_to_unapproved()
                    else:
                        self.load_unapproved_list()
                else:
                    QMessageBox.warning(self, "Nie Znaleziono", "Nie znaleziono plików do usunięcia.")
            except Exception as e:
                QMessageBox.critical(self, "Usuwanie Nie Powiodło Się", f"Nie można usunąć plików:\n{e}")

    def reprocess_current_report(self):
        """Reprocess the current report through AI with preprocessing."""
        if not self.unapproved_reports:
            return
        
        # Use the currently displayed image
        if not self.current_image_path or not os.path.exists(self.current_image_path):
            QMessageBox.warning(self, "Nie Znaleziono", "Nie znaleziono zdjęcia dla bieżącego raportu")
            return
        
        image_path = self.current_image_path
        excel_path = self.unapproved_reports[self.current_report_index]
        
        # Store paths for cleanup after processing
        self.reprocess_old_excel = excel_path
        self.reprocess_old_image = self.find_linked_image(excel_path)
        
        # Close current view of the Excel file
        if self.excel_handler.current_workbook:
            self.excel_handler.current_workbook.close()
        self.excel_handler.current_workbook = None
        self.excel_handler.current_df = None
        self.excel_handler.file_path = None
        
        # Show progress dialog
        self.reprocess_progress = QProgressDialog("Przetwarzanie przez AI...", "Anuluj", 0, 0, self)
        self.reprocess_progress.setWindowModality(Qt.NonModal)
        self.reprocess_progress.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.reprocess_progress.show()
        
        # Start worker thread
        self.reprocess_worker = ReprocessWorker(image_path, config.REPORTS_ROOT)
        self.reprocess_worker.finished.connect(self.on_reprocess_finished)
        self.reprocess_worker.start()
    
    def on_reprocess_finished(self, success, message, output_path, highlighted):
        """Handle reprocessing completion."""
        self.reprocess_progress.close()
        
        if success:
            print(f"\n=== REPROCESS FINISHED ===")
            print(f"New file created: {output_path}")
            print(f"Old Excel: {self.reprocess_old_excel}")
            print(f"Old Image: {self.reprocess_old_image}")
            
            # Check if result is UNKNOWN
            if "UNKNOWN" in os.path.basename(output_path):
                msg = f"⚠️ AI nie mógł wyodrębnić nazwy firmy (ODBIORCA) ze zdjęcia.\n\n"
                msg += f"Wynik: {os.path.basename(output_path)}\n\n"
                msg += "Możliwe powody:\n"
                msg += "• Niska jakość zdjęcia\n"
                msg += "• Dokument jest niepoprawnie obrócony\n"
                msg += "• Komórka ODBIORCA nie jest wyraźnie widoczna\n\n"
                msg += "Spróbuj najpierw obrócić zdjęcie na stronie Przekształcanie."
                QMessageBox.warning(self, "Problem z Ekstrakcją", msg)
            else:
                QMessageBox.information(self, "Przetwarzanie Zakończone", message)
            
            # Delete old files AFTER showing the message
            try:
                if os.path.exists(self.reprocess_old_excel):
                    os.remove(self.reprocess_old_excel)
                    print(f"Deleted old Excel: {self.reprocess_old_excel}")
            except Exception as e:
                print(f"Warning: Could not delete old Excel: {e}")
            
            try:
                if self.reprocess_old_image and os.path.exists(self.reprocess_old_image):
                    os.remove(self.reprocess_old_image)
                    print(f"Deleted old image: {self.reprocess_old_image}")
            except Exception as e:
                print(f"Warning: Could not delete old image: {e}")
            
            # Refresh list - the new file should now be there
            old_list_count = len(self.unapproved_reports)
            self.unapproved_reports = self.file_manager.get_unapproved_reports()
            print(f"List refreshed: {old_list_count} -> {len(self.unapproved_reports)} files")
            print(f"Files in list: {[os.path.basename(p) for p in self.unapproved_reports]}")
            
            if not self.unapproved_reports:
                self.file_name_label.setText("Nie załadowano pliku")
                self.image_label.setText("Brak zdjęcia")
                print("=== No unapproved reports found ===\n")
                return
            
            # Try to find the newly created file by matching the output_path exactly
            new_file_found = False
            for idx, path in enumerate(self.unapproved_reports):
                # Check if this is the exact file or has the same basename
                if path == output_path or os.path.basename(path) == os.path.basename(output_path):
                    self.current_report_index = idx
                    new_file_found = True
                    print(f"Found new file at index {idx}: {os.path.basename(path)}")
                    break
            
            if not new_file_found:
                # If exact match not found, try partial basename match (without extension)
                new_basename = os.path.basename(output_path).replace('.xlsx', '').replace('.xlsm', '').replace('.xls', '')
                for idx, path in enumerate(self.unapproved_reports):
                    if new_basename in os.path.basename(path):
                        self.current_report_index = idx
                        new_file_found = True
                        print(f"Found new file by basename match at index {idx}: {os.path.basename(path)}")
                        break
            
            if not new_file_found:
                # If still not found, stay at current index (capped to list size)
                if self.current_report_index >= len(self.unapproved_reports):
                    self.current_report_index = len(self.unapproved_reports) - 1
                print(f"New file not found in list, loading index {self.current_report_index}")
            
            print(f"=== Loading report at index {self.current_report_index} ===\n")
            self.load_current_report()
        else:
            QMessageBox.critical(self, "Przetwarzanie Nie Powiodło Się", message)

    def refresh(self): 
        if self.is_review_mode:
            if self.approved_reports and self.current_approved_index < len(self.approved_reports):
                self.load_approved_report(self.approved_reports[self.current_approved_index])
        else:
            self.load_current_report()
    
    def prev_report(self):
        if self.is_review_mode:
            if self.current_approved_index > 0:
                self.current_approved_index -= 1
                self.load_approved_report(self.approved_reports[self.current_approved_index])
            return
        if self.current_report_index > 0:
            self.current_report_index -= 1
            self.load_current_report()

    def next_report(self):
        if self.is_review_mode:
            if self.current_approved_index < len(self.approved_reports) - 1:
                self.current_approved_index += 1
                self.load_approved_report(self.approved_reports[self.current_approved_index])
            return
        if self.current_report_index < len(self.unapproved_reports) - 1:
            self.current_report_index += 1
            self.load_current_report()
            
    def show_full_image(self):
        if self.current_image_path and os.path.exists(self.current_image_path):
             os.startfile(self.current_image_path)
    
    def open_excel_file(self):
        if self.current_excel_path and os.path.exists(self.current_excel_path):
            try:
                os.startfile(self.current_excel_path)
            except Exception as e:
                QMessageBox.warning(self, "Błąd", f"Nie można otworzyć pliku:\n{str(e)}")
        else:
            QMessageBox.information(self, "Brak Pliku", "Nie załadowano żadnego pliku Excel.")
    
    def reload_approved_list(self):
        """Reload the approved reports list from database"""
        try:
            from core.database_handler import DatabaseHandler
            from config import DATABASE_FILE
            
            db = DatabaseHandler(DATABASE_FILE)
            records = db.get_all_approved_records()
            
            # Extract filepaths from records (only files that still exist)
            self.approved_reports = [record['filepath'] for record in records if os.path.exists(record['filepath'])]
            
        except Exception as e:
            print(f"Error reloading approved list: {e}")
            self.approved_reports = []
    
    def back_to_unapproved(self):
        self.is_review_mode = False
        self.approved_reports = []
        self.current_approved_index = 0
        self.back_to_unapproved_btn.hide()
        self.select_unapproved_btn.show()
        self.approve_btn.show()
        self.reprocess_btn.show()
        self.load_unapproved_list()
    
    def show_unapproved_dialog(self):
        if not self.unapproved_reports:
            QMessageBox.information(self, "Brak Raportów", "Brak dostępnych niezatwierdzonych raportów.")
            return
        dlg = UnapprovedReportsDialog(self.unapproved_reports, self)
        if dlg.exec_() == QDialog.Accepted:
            path_to_open = dlg.selected_file_path
            if path_to_open and path_to_open in self.unapproved_reports:
                self.current_report_index = self.unapproved_reports.index(path_to_open)
                self.load_current_report()
    
    def load_approved_report(self, path_to_open):
        try:
            if path_to_open is None or not isinstance(path_to_open, str):
                raise Exception(f"Invalid path_to_open: {path_to_open} (type: {type(path_to_open)})")
            
            if not os.path.exists(path_to_open):
                raise Exception(f"File not found at: {path_to_open}")
            
            self.current_excel_path = path_to_open
            
            df = self.excel_handler.load_file(path_to_open)
            style_map = self.excel_handler.get_formatting()
            
            filename = os.path.basename(path_to_open)
            
            self.file_icon_label.setText("✓")
            self._full_file_name_text = f"Excel: {filename} (Approved)"
            QTimer.singleShot(0, self.update_file_name_label_display)
            
            self.table.blockSignals(True)
            self.table.clear()
            self.table.setRowCount(len(df))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels([f"Col {i+1}" for i in range(len(df.columns))])

            for r in range(len(df)):
                for c in range(len(df.columns)):
                    cell_value = df.iloc[r, c]
                    if pd.isna(cell_value) or cell_value is None:
                        val = ""
                    else:
                        val = str(cell_value)
                    item = QTableWidgetItem(val)
                    if (r, c) in style_map:
                        style = style_map[(r, c)]
                        if style['bg']: item.setBackground(QBrush(QColor(style['bg'])))
                        if style['fg']: item.setForeground(QBrush(QColor(style['fg'])))
                    self.table.setItem(r, c, item)
            
            self.table.resizeColumnsToContents()

            # Cap widths for approved view as well
            if self.table.columnCount() > 1:
                self.table.setColumnWidth(1, 250)
            for c in range(self.table.columnCount()):
                if c == 1:
                    continue
                current_width = self.table.columnWidth(c)
                new_width = min(current_width + 8, 220)
                self.table.setColumnWidth(c, new_width)
            
            self.table.blockSignals(False)
            # Apply default 60:40 split after render unless user adjusted manually
            QTimer.singleShot(0, self.apply_default_split)
            QTimer.singleShot(0, self.clamp_table_width)

            # Load linked image
            self.current_image_path = self.find_linked_image(path_to_open)
            if self.current_image_path and os.path.exists(self.current_image_path):
                pic_name = os.path.basename(self.current_image_path)
                self.picture_name_label.setText(f"🖼️ Zdjęcie: {pic_name}")
                pix = QPixmap(self.current_image_path)
                self.image_label.setPixmap(pix)
            else:
                self.picture_name_label.setText(f"🖼️ Brak zdjęcia")
                self.image_label.setText("Brak dostępnego zdjęcia")
                self.image_label.setPixmap(None)

            self.nav_label.setText(f"ZATWIERDZONE {self.current_approved_index + 1}/{len(self.approved_reports)}")
            
        except Exception as e:
            QMessageBox.critical(self, "Błąd Ładowania", f"Nie można otworzyć raportu:\n{e}")
    
    def show_approved_dialog(self):
        # Get current report's month if available
        filter_month = None
        try:
            if self.excel_handler.current_workbook:
                ws = self.excel_handler.current_workbook.active
                date_cell = ws['D1'].value
                if date_cell:
                    from datetime import datetime
                    if isinstance(date_cell, datetime):
                        filter_month = date_cell.strftime('%Y-%m')
                    elif isinstance(date_cell, str):
                        try:
                            parsed_date = datetime.strptime(date_cell.strip(), '%d.%m.%Y')
                            filter_month = parsed_date.strftime('%Y-%m')
                        except:
                            pass
        except:
            pass
        
        dlg = ApprovedReportsDialog(self, filter_month)
        if dlg.exec_() == QDialog.Accepted:
            path_to_open = dlg.selected_file_path
            
            if path_to_open and isinstance(path_to_open, str) and len(path_to_open) > 0:
                # Build approved_reports list from all approved records for navigation
                try:
                    from core.database_handler import DatabaseHandler
                    from config import DATABASE_FILE
                    db = DatabaseHandler(DATABASE_FILE)
                    records = db.get_all_approved_records()
                    all_filepaths = [record['filepath'] for record in records if os.path.exists(record['filepath'])]
                    self.approved_reports = all_filepaths
                except Exception as e:
                    print(f"Error loading approved reports: {e}")
                    self.approved_reports = [path_to_open]
                
                # Set current index if the selected file is in the list
                if path_to_open in self.approved_reports:
                    self.current_approved_index = self.approved_reports.index(path_to_open)
                else:
                    # If not in list, add it and use it as current
                    self.approved_reports = [path_to_open]
                    self.current_approved_index = 0
                
                self.show_normal_view()
                self.is_review_mode = True
                self.select_unapproved_btn.hide()
                self.back_to_unapproved_btn.show()
                self.approve_btn.hide()
                self.reprocess_btn.hide()
                self.load_approved_report(path_to_open)



# --- 3. The Main App (Tabs) ---
class VerifyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("System Weryfikacji Excel")
        self.resize(1200, 800)

        # APPLY THE GLOBAL STYLESHEET HERE
        self.setStyleSheet(STYLESHEET)
        
        # Main Layout
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Tabs
        self.tabs = QTabWidget()
        self.tabs.tabBar().setExpanding(False)
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #E5E7EB;
                margin: 0px;
            }
            QTabBar::tab {
                padding: 10px 15px;
                margin-right: 2px;
                margin-top: 0px;
                margin-bottom: 0px;
                font-size: 14px;
                min-width: 130px;
            }
        """)
        
        # Settings button
        settings_btn = QToolButton()
        settings_btn.setText("\u2699")
        settings_btn.setFixedSize(44, 32)
        settings_btn.setCursor(Qt.PointingHandCursor)
        settings_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
        settings_btn.setStyleSheet("""
            QToolButton {
                background-color: #F3F4F6;
                border: 1px solid #D1D5DB;
                border-radius: 6px;
                font-size: 20px;
                padding: 0px;
                margin: 0px 4px 0px 0px;
                line-height: 34px;
            }
            QToolButton:hover {
                background-color: #E5E7EB;
                border-color: #9CA3AF;
            }
        """)
        settings_btn.setContentsMargins(0, 0, 0, 0)
        settings_btn.clicked.connect(self.open_settings)
        
        # Import/Export button
        import_export_btn = QToolButton()
        import_export_btn.setText("📦")
        import_export_btn.setFixedSize(44, 32)
        import_export_btn.setCursor(Qt.PointingHandCursor)
        import_export_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
        import_export_btn.setToolTip("Import / Export Danych")
        import_export_btn.setStyleSheet("""
            QToolButton {
                background-color: white;
                border: 1px solid #D1D5DB;
                border-radius: 6px;
                font-size: 20px;
                padding: 0px;
                margin: 0px 4px 0px 0px;
                line-height: 34px;
            }
            QToolButton:hover {
                background-color: #E5E7EB;
                border-color: #9CA3AF;
            }
        """)
        import_export_btn.setContentsMargins(0, 0, 0, 0)
        import_export_btn.clicked.connect(self.open_import_export)
        
        # Container for corner buttons
        corner_widget = QWidget()
        corner_layout = QHBoxLayout(corner_widget)
        corner_layout.setContentsMargins(0, 0, 10, 0)
        corner_layout.setSpacing(4)
        corner_layout.addWidget(import_export_btn)
        corner_layout.addWidget(settings_btn)
        
        self.tabs.setCornerWidget(corner_widget, Qt.TopRightCorner)
        
        layout.addWidget(self.tabs)
        
        # Create Instances of pages
        self.tab1 = VerificationPage()
        self.tab2 = TransformPage()
        self.tab3 = GenerateReportPage()
        
        # Connect transformation complete signal to refresh verification page
        self.tab2.transformation_complete.connect(self.tab1.load_unapproved_list)
        
        # Connect tab change to refresh GenerateReportPage months
        self.tabs.currentChanged.connect(self.on_tab_changed)
        
        # Add tabs
        self.tabs.addTab(self.tab1, "Weryfikacja")
        self.tabs.addTab(self.tab2, "Zdjęcie na Excel")
        self.tabs.addTab(self.tab3, "Generuj Raport")
   
    def on_tab_changed(self, index):
        """Refresh data when switching to specific tabs."""
        # Refresh GenerateReportPage when switching to it (index 2)
        if index == 2:
            self.tab3.refresh_months()
    
    def open_settings(self):
        dlg = SettingsDialog(self)
        dlg.exec_()
    
    def open_import_export(self):
        """Open import/export dialog."""
        dlg = ImportExportDialog(self)
        dlg.data_refreshed.connect(self.refresh_all_lists)
        dlg.exec_()
    
    def refresh_all_lists(self):
        """Refresh both unapproved and approved report lists after import."""
        try:
            self.tab1.load_unapproved_list()
            self.tab1.reload_approved_list()
        except Exception as e:
            print(f"Error refreshing lists: {e}")

    def _create_styled_button(self, text, width):
        btn = QPushButton(text)
        btn.setFixedWidth(width)
        btn.setFixedHeight(45)
        btn.setCursor(Qt.PointingHandCursor)
        return btn