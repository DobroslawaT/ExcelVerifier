import os

import pandas as pd
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QFileDialog, QMessageBox, QComboBox
)
from PyQt5.QtCore import Qt

import config
from core.company_db import load_company_db, save_company_db, normalize_nip, merge_companies


class CompanyDbDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Dodaj / usuń firmę")
        self.resize(720, 520)

        if parent:
            self.setStyleSheet(parent.styleSheet())

        self.db_file = config.COMPANY_DB_FILE
        try:
            loaded = load_company_db(self.db_file)
            self.companies = loaded if loaded else []
        except Exception as e:
            print(f"Error loading company database: {e}")
            self.companies = []
        
        # Clean up - remove companies without valid NIPs
        cleaned_companies = []
        for c in self.companies:
            try:
                if isinstance(c, dict):
                    nip = str(c.get("nip", "")).strip()
                    if nip:  # Only keep companies with non-empty NIP
                        c["nip"] = nip
                        cleaned_companies.append(c)
            except Exception as e:
                print(f"Error processing company: {e}")
                continue
        
        self.companies = cleaned_companies
        self.filtered_companies = list(self.companies)

        self.init_ui()
        self.populate_table()

    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        header = QLabel("Baza firm do weryfikacji (nazwa + NIP)")
        header.setStyleSheet("font-size: 16px; font-weight: bold; color: #111827;")
        layout.addWidget(header)

        search_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Szukaj po firmie lub NIP...")
        self.search_input.textChanged.connect(self.apply_filter)
        self.search_input.installEventFilter(self)
        search_row.addWidget(self.search_input, 1)
        layout.addLayout(search_row)

        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Firma", "NIP"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        layout.addWidget(self.table, 1)

        actions_row = QHBoxLayout()
        add_btn = QPushButton("Dodaj firmę")
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.setAutoDefault(False)
        add_btn.setDefault(False)
        add_btn.clicked.connect(self.open_add_dialog)

        import_btn = QPushButton("Importuj z Excela")
        import_btn.setCursor(Qt.PointingHandCursor)
        import_btn.setAutoDefault(False)
        import_btn.setDefault(False)
        import_btn.clicked.connect(self.import_from_excel)

        edit_btn = QPushButton("Edytuj zaznaczoną")
        edit_btn.setCursor(Qt.PointingHandCursor)
        edit_btn.setAutoDefault(False)
        edit_btn.setDefault(False)
        edit_btn.clicked.connect(self.open_edit_dialog)

        delete_btn = QPushButton("Usuń zaznaczoną")
        delete_btn.setCursor(Qt.PointingHandCursor)
        delete_btn.setAutoDefault(False)
        delete_btn.setDefault(False)
        delete_btn.clicked.connect(self.delete_selected)

        close_btn = QPushButton("Zamknij")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setAutoDefault(False)
        close_btn.setDefault(False)
        close_btn.clicked.connect(self.accept)

        actions_row.addWidget(add_btn)
        actions_row.addWidget(import_btn)
        actions_row.addWidget(edit_btn)
        actions_row.addStretch(1)
        actions_row.addWidget(delete_btn)
        actions_row.addWidget(close_btn)
        layout.addLayout(actions_row)

    def populate_table(self, items=None):
        if items is None:
            items = self.filtered_companies
        
        self.table.setRowCount(0)  # Clear first
        
        for row_idx, item in enumerate(items):
            try:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("name", "")).strip()
                nip = str(item.get("nip", "")).strip()
                
                self.table.insertRow(row_idx)
                self.table.setItem(row_idx, 0, QTableWidgetItem(name))
                self.table.setItem(row_idx, 1, QTableWidgetItem(nip))
            except Exception as e:
                print(f"Error populating row {row_idx}: {e}")
                continue

    def open_add_dialog(self):
        dialog = AddCompanyDialog(self)
        if dialog.exec_() != QDialog.Accepted:
            return
        name, nip = dialog.get_values()
        self.add_company_data(name, nip)

    def open_edit_dialog(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Brak wyboru", "Wybierz firmę do edycji.")
            return

        current_name = self.table.item(row, 0).text()
        current_nip = self.table.item(row, 1).text()

        if not current_nip or current_nip.strip() == "":
            QMessageBox.warning(self, "Brak NIP", "Nie można edytować firmy bez NIP. Usuń ją i dodaj ponownie z NIP.")
            return

        dialog = AddCompanyDialog(self, title="Edytuj firmę")
        dialog.set_values(current_name, current_nip)
        if dialog.exec_() != QDialog.Accepted:
            return
        new_name, new_nip = dialog.get_values()
        self.edit_company_data(current_nip, new_name, new_nip)

    def add_company_data(self, name, nip):
        print(f"[ADD] Starting with name={name}, nip={nip}")
        name = name.strip()
        nip = normalize_nip(nip)
        print(f"[ADD] After normalize: name={name}, nip={nip}")

        if not name or not nip:
            QMessageBox.warning(self, "Brak danych", "Podaj nazwę firmy i NIP.")
            return
        if len(nip) != 10:
            QMessageBox.warning(self, "Nieprawidłowy NIP", "NIP powinien mieć 10 cyfr.")
            return

        print(f"[ADD] Current companies before: {self.companies}")
        existing = next((c for c in self.companies if c.get("nip") == nip), None)
        if existing:
            existing["name"] = name
            message = "Zaktualizowano nazwę firmy dla istniejącego NIP."
            print(f"[ADD] Updated existing company")
        else:
            self.companies.append({"name": name, "nip": nip})
            message = "Dodano firmę do bazy."
            print(f"[ADD] Added new company")

        print(f"[ADD] Current companies after: {self.companies}")
        if not save_company_db(self.db_file, self.companies):
            QMessageBox.critical(self, "Błąd", "Nie udało się zapisać bazy firm.")
            print(f"[ADD] Save failed!")
            return

        print(f"[ADD] Save succeeded, reloading...")
        # Reload from database to get fresh data
        self.reload_companies()
        print(f"[ADD] After reload: {self.companies}")
        self.apply_filter()
        QMessageBox.information(self, "Gotowe", message)

    def edit_company_data(self, original_nip, name, nip):
        name = name.strip()
        nip = normalize_nip(nip)

        if not name or not nip:
            QMessageBox.warning(self, "Brak danych", "Podaj nazwę firmy i NIP.")
            return
        if len(nip) != 10:
            QMessageBox.warning(self, "Nieprawidłowy NIP", "NIP powinien mieć 10 cyfr.")
            return

        if nip != original_nip:
            conflict = next((c for c in self.companies if c.get("nip") == nip), None)
            if conflict:
                QMessageBox.warning(self, "Konflikt", "Istnieje już firma z podanym NIP.")
                return

        updated = False
        for item in self.companies:
            if item.get("nip") == original_nip:
                item["name"] = name
                item["nip"] = nip
                updated = True
                break

        if not updated:
            self.companies.append({"name": name, "nip": nip})

        if not save_company_db(self.db_file, self.companies):
            QMessageBox.critical(self, "Błąd", "Nie udało się zapisać bazy firm.")
            return

        # Reload from database to get fresh data
        self.reload_companies()
        self.apply_filter()
        QMessageBox.information(self, "Gotowe", "Zaktualizowano dane firmy.")

    def apply_filter(self):
        text = self.search_input.text().strip().lower()
        if not text:
            self.filtered_companies = list(self.companies)
        else:
            self.filtered_companies = [
                item for item in self.companies
                if text in item.get("name", "").lower() or text in item.get("nip", "")
            ]
        self.populate_table()

    def reload_companies(self):
        """Reload companies from database and clean them up (remove those without NIPs)."""
        print(f"[RELOAD] Starting reload...")
        try:
            loaded = load_company_db(self.db_file)
            self.companies = loaded if loaded else []
            print(f"[RELOAD] Loaded {len(self.companies)} companies from database")
        except Exception as e:
            print(f"[RELOAD] Error loading: {e}")
            import traceback
            traceback.print_exc()
            self.companies = []
        
        # Clean up - remove companies without valid NIPs
        cleaned_companies = []
        for c in self.companies:
            try:
                if isinstance(c, dict):
                    nip = str(c.get("nip", "")).strip()
                    if nip:  # Only keep companies with non-empty NIP
                        c["nip"] = nip
                        cleaned_companies.append(c)
            except Exception as e:
                print(f"[RELOAD] Error processing company: {e}")
                continue
        
        print(f"[RELOAD] After cleanup: {len(cleaned_companies)} companies")
        self.companies = cleaned_companies
        self.filtered_companies = list(self.companies)
        print(f"[RELOAD] Final companies list: {self.companies}")

    def eventFilter(self, source, event):
        if source is self.search_input and event.type() == event.KeyPress:
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                return True
        return super().eventFilter(source, event)

    def delete_selected(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Brak wyboru", "Wybierz firmę do usunięcia.")
            return

        name = self.table.item(row, 0).text()
        nip = self.table.item(row, 1).text()
        confirm = QMessageBox.question(
            self,
            "Usuń firmę",
            f"Czy na pewno chcesz usunąć firmę:\n\n{name} (NIP: {nip})",
            QMessageBox.Yes | QMessageBox.No
        )
        if confirm != QMessageBox.Yes:
            return

        self.companies = [c for c in self.companies if c.get("nip") != nip]
        if not save_company_db(self.db_file, self.companies):
            QMessageBox.critical(self, "Błąd", "Nie udało się zapisać bazy firm.")
            return

        self.apply_filter()

    def import_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Wybierz plik Excel",
            os.getcwd(),
            "Excel Files (*.xlsx *.xls)"
        )
        if not file_path:
            return

        try:
            workbook = pd.ExcelFile(file_path, engine="openpyxl")
        except Exception as exc:
            QMessageBox.critical(self, "Błąd", f"Nie udało się wczytać pliku: {exc}")
            return

        sheet_picker = SheetPickerDialog(self, workbook.sheet_names)
        if sheet_picker.exec_() != QDialog.Accepted:
            return
        sheet_name = sheet_picker.get_selection()

        try:
            df = pd.read_excel(workbook, sheet_name=sheet_name)
        except Exception as exc:
            QMessageBox.critical(self, "Błąd", f"Nie udało się wczytać arkusza: {exc}")
            return

        if df.empty:
            QMessageBox.warning(self, "Brak danych", "Wybrany plik jest pusty.")
            return

        name_col = self._pick_column(df, {"firma", "company", "nazwa", "nazwa firmy", "kontrahent"})
        nip_col = self._pick_column(df, {"nip"})
        mapping = ColumnMappingDialog(self, df, name_col, nip_col)
        if mapping.exec_() != QDialog.Accepted:
            return
        name_col, nip_col = mapping.get_selection()
        if name_col is None or nip_col is None:
            QMessageBox.warning(self, "Brak kolumn", "Wybierz kolumny dla nazwy firmy i NIP.")
            return

        new_items = []
        skipped_empty = 0
        skipped_invalid = 0
        total_rows = len(df)
        seen_nips = set()
        duplicate_in_import = 0
        for _, row in df.iterrows():
            name = str(row.get(name_col, "")).strip()
            nip = normalize_nip(row.get(nip_col, ""))
            if not name or not nip:
                skipped_empty += 1
                continue
            if len(nip) != 10:
                skipped_invalid += 1
                continue
            if nip in seen_nips:
                duplicate_in_import += 1
                continue
            seen_nips.add(nip)
            new_items.append({"name": name, "nip": nip})

        if not new_items:
            QMessageBox.warning(self, "Brak danych", "Nie znaleziono poprawnych wpisów do importu.")
            return

        existing_by_nip = {item.get("nip"): item for item in self.companies}
        updated_count = 0
        for item in new_items:
            existing = existing_by_nip.get(item.get("nip"))
            if existing and existing.get("name") != item.get("name"):
                updated_count += 1

        merged = merge_companies(self.companies, new_items)
        if not save_company_db(self.db_file, merged):
            QMessageBox.critical(self, "Błąd", "Nie udało się zapisać bazy firm.")
            return

        added_count = len(merged) - len(self.companies)
        self.companies = merged
        self.apply_filter()

        QMessageBox.information(
            self,
            "Import zakończony",
            "\n".join([
                f"Arkusz: {sheet_name}",
                f"Wierszy w arkuszu: {total_rows}",
                f"Nowych firm: {added_count}",
                f"Zaktualizowano nazwy: {updated_count}",
                f"Duplikaty w imporcie: {duplicate_in_import}",
                f"Pominieto (brak danych): {skipped_empty}",
                f"Pominieto (zly NIP): {skipped_invalid}",
            ])
        )

    def _pick_column(self, df, candidates):
        for col in df.columns:
            if str(col).strip().lower() in candidates:
                return col
        return None


class ColumnMappingDialog(QDialog):
    def __init__(self, parent, df, suggested_name, suggested_nip):
        super().__init__(parent)
        self.setWindowTitle("Wybierz kolumny")
        self.resize(480, 220)
        self.df = df
        self.name_col = suggested_name
        self.nip_col = suggested_nip

        if parent:
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        hint = QLabel("Wybierz kolumny dla nazwy firmy i NIP:")
        hint.setStyleSheet("font-size: 13px; color: #374151;")
        layout.addWidget(hint)

        col_names = [str(col) for col in self.df.columns]
        if not col_names:
            col_names = [""]

        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Nazwa firmy"))
        self.name_combo = QComboBox()
        self.name_combo.addItems(col_names)
        name_row.addWidget(self.name_combo)
        layout.addLayout(name_row)

        nip_row = QHBoxLayout()
        nip_row.addWidget(QLabel("NIP"))
        self.nip_combo = QComboBox()
        self.nip_combo.addItems(col_names)
        nip_row.addWidget(self.nip_combo)
        layout.addLayout(nip_row)

        if self.name_col in self.df.columns:
            self.name_combo.setCurrentText(str(self.name_col))
        if self.nip_col in self.df.columns:
            self.nip_combo.setCurrentText(str(self.nip_col))

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel_btn = QPushButton("Anuluj")
        cancel_btn.clicked.connect(self.reject)
        ok_btn = QPushButton("Importuj")
        ok_btn.clicked.connect(self.accept)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(ok_btn)
        layout.addLayout(buttons)

    def get_selection(self):
        return self.name_combo.currentText(), self.nip_combo.currentText()


class AddCompanyDialog(QDialog):
    def __init__(self, parent=None, title="Dodaj firmę"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(680, 240)

        if parent:
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        hint = QLabel("Podaj nazwę firmy i NIP:")
        hint.setStyleSheet("font-size: 13px; color: #374151;")
        layout.addWidget(hint)

        name_label = QLabel("Nazwa firmy")
        layout.addWidget(name_label)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Nazwa firmy")
        self.name_input.setMinimumHeight(36)
        layout.addWidget(self.name_input)

        nip_label = QLabel("NIP")
        layout.addWidget(nip_label)
        self.nip_input = QLineEdit()
        self.nip_input.setPlaceholderText("NIP (10 cyfr)")
        self.nip_input.setMaxLength(20)
        self.nip_input.setMinimumHeight(36)
        layout.addWidget(self.nip_input)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel_btn = QPushButton("Anuluj")
        cancel_btn.clicked.connect(self.reject)
        ok_btn = QPushButton("Dodaj")
        ok_btn.clicked.connect(self.accept)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(ok_btn)
        layout.addLayout(buttons)

    def get_values(self):
        return self.name_input.text(), self.nip_input.text()

    def set_values(self, name, nip):
        self.name_input.setText(name)
        self.nip_input.setText(nip)

    def accept(self):
        name = self.name_input.text().strip()
        nip = normalize_nip(self.nip_input.text())
        if not name or not nip:
            QMessageBox.warning(self, "Brak danych", "Podaj nazwę firmy i NIP.")
            return
        if len(nip) != 10:
            QMessageBox.warning(self, "Nieprawidłowy NIP", "NIP powinien mieć 10 cyfr.")
            return
        super().accept()


class SheetPickerDialog(QDialog):
    def __init__(self, parent, sheet_names):
        super().__init__(parent)
        self.setWindowTitle("Wybierz arkusz")
        self.resize(360, 160)
        self.sheet_names = sheet_names or [""]

        if parent:
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        label = QLabel("Wybierz arkusz do importu:")
        label.setStyleSheet("font-size: 13px; color: #374151;")
        layout.addWidget(label)

        self.sheet_combo = QComboBox()
        self.sheet_combo.addItems([str(name) for name in self.sheet_names])
        layout.addWidget(self.sheet_combo)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel_btn = QPushButton("Anuluj")
        cancel_btn.clicked.connect(self.reject)
        ok_btn = QPushButton("Dalej")
        ok_btn.clicked.connect(self.accept)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(ok_btn)
        layout.addLayout(buttons)

    def get_selection(self):
        return self.sheet_combo.currentText()
