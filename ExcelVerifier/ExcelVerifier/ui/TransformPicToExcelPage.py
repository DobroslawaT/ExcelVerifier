import os
import glob
import shutil
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QLabel, QListWidget, 
    QListWidgetItem, QMessageBox, QHBoxLayout, QFrame, 
    QProgressDialog, QStyle, QSplitter, QFileDialog, QSizePolicy, QAbstractItemView
)
from PyQt5.QtGui import QIcon, QPixmap, QTransform, QPainter, QColor, QPen, QPainterPath
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QRectF
from PIL import Image, ImageChops

# Ensure you have these modules in your project
from core.image_transformer import ImageTransformer
import config


def trim_whitespace(path: str, out_path: str = None, bg_threshold: int = 250, margin: int = 0):
    """
    Trim white borders from an image (optimized for speed).
    Returns the path to the trimmed image.
    """
    img = Image.open(path).convert("RGB")
    
    # OPTIMIZATION: Work on smaller version for speed (4x faster)
    max_size = 2000
    original_size = img.size
    if max(img.size) > max_size:
        scale = max_size / max(img.size)
        new_size = (int(img.width * scale), int(img.height * scale))
        img_small = img.resize(new_size, Image.LANCZOS)
    else:
        img_small = img
        scale = 1.0
    
    # Detect white borders on smaller image
    bg = Image.new("RGB", img_small.size, (255, 255, 255))
    diff = ImageChops.difference(img_small, bg)
    diff = ImageChops.add(diff, diff, 2.0, 0)
    diff = diff.point(lambda p: 255 if p > (255 - bg_threshold) else 0)
    bbox = diff.getbbox()
    
    if not bbox:
        # Image appears fully white; return original
        return path
    
    # Scale bounding box back to original size
    if scale != 1.0:
        left, top, right, bottom = bbox
        bbox = (
            int(left / scale),
            int(top / scale),
            int(right / scale),
            int(bottom / scale)
        )
    
    if margin:
        left, top, right, bottom = bbox
        left = max(left - margin, 0)
        top = max(top - margin, 0)
        right = min(right + margin, original_size[0])
        bottom = min(bottom + margin, original_size[1])
        bbox = (left, top, right, bottom)
    
    cropped = img.crop(bbox)
    out = out_path or str(Path(path).with_name(f"{Path(path).stem}_trimmed{Path(path).suffix}"))
    cropped.save(out)
    return out

# --- 1. Custom Icon Painter (Sharp & Proportional) ---
def create_painted_icon(name, color_hex="#374151", size=64):
    """Draws sharp, proportional vector icons."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    
    pen = QPen(QColor(color_hex))
    pen.setWidthF(2.5)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)

    c = size / 2.0
    painter.translate(c, c)
    path = QPainterPath()

    if name == "trash":
        # Proportional Trash Can
        body_w, body_h, lid_w, lid_y = 10, 12, 14, -8
        path.moveTo(-lid_w, lid_y); path.lineTo(lid_w, lid_y)
        path.moveTo(-body_w, lid_y); path.lineTo(-body_w*0.9, body_h); path.lineTo(body_w*0.9, body_h); path.lineTo(body_w, lid_y)
        path.moveTo(-3, lid_y); path.lineTo(-3, lid_y-4); path.lineTo(3, lid_y-4); path.lineTo(3, lid_y)
    elif name == "rotate_left":
        r = 10
        path.arcMoveTo(-r, -r, 2*r, 2*r, 90); path.arcTo(QRectF(-r, -r, 2*r, 2*r), 90, 270)
        path.moveTo(0, -r); path.lineTo(-5, -r-3); path.moveTo(0, -r); path.lineTo(-1, -r+5)
    elif name == "rotate_right":
        r = 10
        path.arcMoveTo(-r, -r, 2*r, 2*r, 90); path.arcTo(QRectF(-r, -r, 2*r, 2*r), 90, -270)
        path.moveTo(0, -r); path.lineTo(5, -r-3); path.moveTo(0, -r); path.lineTo(1, -r+5)

    painter.drawPath(path)
    painter.end()
    return QIcon(pixmap)

# --- 2. Custom List Item Widget (The "Podgląd" Row - Clean White) ---
class FileListItemWidget(QWidget):
    """
    Custom row: [ Thumbnail ] [ Filename ] [ Spacer ] [ X Button ]
    """
    def __init__(self, path, on_delete_clicked):
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(10)
        
        # *** FIX: Set background to transparent to remove the gray box ***
        self.setStyleSheet("background-color: transparent;")
        
        # 1. Thumbnail (Podgląd)
        self.thumb = QLabel()
        self.thumb.setFixedSize(40, 40)
        # *** FIX: Set background to white (not gray) ***
        self.thumb.setStyleSheet("border: 1px solid #E2E8F0; border-radius: 4px; background-color: white;")
        self.thumb.setAlignment(Qt.AlignCenter)
        
        pix = QPixmap(path)
        if not pix.isNull():
            self.thumb.setPixmap(pix.scaled(38, 38, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            self.thumb.setText("?")
        layout.addWidget(self.thumb)
        
        # 2. Filename
        name_lbl = QLabel(os.path.basename(path))
        # Ensure text has no background
        name_lbl.setStyleSheet("font-size: 13px; color: #334155; font-weight: 500; background-color: transparent;")
        layout.addWidget(name_lbl)
        
        # 3. Spacer
        layout.addStretch()
        
        # 4. Delete Button (X)
        del_btn = QPushButton("✕")
        del_btn.setFixedSize(24, 24)
        del_btn.setCursor(Qt.PointingHandCursor)
        del_btn.setToolTip("Remove from list")
        del_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent; color: #94A3B8; border: none; font-weight: bold; font-size: 12px;
                border-radius: 12px;
            }
            QPushButton:hover {
                background-color: #FEF2F2; color: #DC2626;
            }
        """)
        del_btn.clicked.connect(on_delete_clicked)
        layout.addWidget(del_btn)

# --- 3. Scalable Label (For Main Preview) ---
class ScalableImageLabel(QLabel):
    def __init__(self):
        super().__init__()
        self.setAlignment(Qt.AlignCenter)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.setMinimumSize(100, 100)
        self._pixmap = None

    def set_image(self, pixmap):
        self._pixmap = pixmap
        self._update_display()

    def resizeEvent(self, event):
        self._update_display()
        super().resizeEvent(event)

    def _update_display(self):
        if self._pixmap and not self._pixmap.isNull():
            scaled = self._pixmap.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            super().setPixmap(scaled)
        else:
            super().setText("Wybierz zdjęcie do podglądu")

# --- 4. Drop Zone ---
class DropZone(QFrame):
    def __init__(self, on_files_dropped, on_select_dir=None, on_select_files=None):
        super().__init__()
        self.on_files_dropped = on_files_dropped
        self.setAcceptDrops(True)
        self.setStyleSheet("QFrame { border: 2px dashed #CBD5E1; border-radius: 8px; background-color: #F8FAFC; }")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        lbl = QLabel("Przeciągnij zdjęcia lub użyj przycisków poniżej", self)
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet("font-size: 14px; color: #64748B; padding-bottom: 10px; background-color: transparent;")
        layout.addWidget(lbl)
        
        if on_select_dir or on_select_files:
            btn_row = QHBoxLayout()
            btn_row.setSpacing(10)
            btn_row.setAlignment(Qt.AlignCenter)
            def mk_btn(txt, slot):
                b = QPushButton(txt)
                b.setCursor(Qt.PointingHandCursor)
                b.setStyleSheet("QPushButton { background: white; border: 1px solid #CBD5E1; border-radius: 6px; padding: 6px 14px; font-weight: 600; color: #475569; } QPushButton:hover { background: #F1F5F9; border-color: #94A3B8; color: #1E293B; }")
                b.clicked.connect(slot)
                return b
            if on_select_dir: btn_row.addWidget(mk_btn("Wybierz folder", on_select_dir))
            if on_select_files: btn_row.addWidget(mk_btn("Wybierz pliki", on_select_files))
            layout.addLayout(btn_row)

    def dragEnterEvent(self, e): e.acceptProposedAction(); self.setStyleSheet("QFrame { border: 2px dashed #3B82F6; background: #EFF6FF; }")
    def dragLeaveEvent(self, e): self.setStyleSheet("QFrame { border: 2px dashed #CBD5E1; background: #F8FAFC; }")
    def dropEvent(self, e):
        self.setStyleSheet("QFrame { border: 2px dashed #CBD5E1; background: #F8FAFC; }")
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        imgs = [p for p in paths if os.path.splitext(p)[1].lower() in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']]
        if imgs: self.on_files_dropped(imgs)

# --- 5. Worker Thread ---
class TransformWorker(QThread):
    progress, finished, error = pyqtSignal(int, int), pyqtSignal(list), pyqtSignal(str)
    def __init__(self, paths, folder=None, max_workers=3): 
        super().__init__()
        self.paths = paths
        # Use REPORTS_ROOT (Niezatwierdzone) so files appear in verification tab
        self.folder = folder or config.REPORTS_ROOT
        self.max_workers = max_workers  # Process 2-3 images in parallel
    
    def run(self):
        try:
            t = ImageTransformer()
            res = []
            completed = 0
            
            # Use ThreadPoolExecutor for parallel image processing (2-3 concurrent)
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # Submit all tasks
                future_to_path = {
                    executor.submit(t.process_image_file, p, self.folder): p 
                    for p in self.paths
                }
                
                # Collect results as they complete
                for future in as_completed(future_to_path):
                    p = future_to_path[future]
                    completed += 1
                    self.progress.emit(completed, len(self.paths))
                    
                    try:
                        out, h = future.result()
                        res.append((p, True, out, h))
                    except Exception as e:
                        res.append((p, False, str(e), 0))
            
            self.finished.emit(res)
        except Exception as e:
            self.error.emit(str(e))


# --- 6. Main Page ---
class TransformPage(QWidget):
    transformation_complete = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.selected_images = []
        self.image_rotations = {}
        self.current_preview_index = -1
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        # Header
        layout.addWidget(QLabel("Przekształć Zdjęcia w Excel", styleSheet="font-size: 24px; font-weight: 700; color: #0F172A;"))
        layout.addWidget(QLabel("Wybierz zdjęcia do przekształcenia w arkusze kalkulacyjne", styleSheet="font-size: 14px; color: #64748B;"))
        
        # Drop Zone
        self.drop_zone = DropZone(self.handle_files, self.sel_dir, self.sel_files)
        layout.addWidget(self.drop_zone)

        # Status
        status_row = QHBoxLayout()
        self.status_label = QLabel("Nie załadowano zdjęć", styleSheet="color: #64748B; font-weight: 500;")
        status_row.addWidget(self.status_label, 1)
        self.clear_btn = QPushButton("Wyczyść wszystko")
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.setStyleSheet("QPushButton { background: white; color: #DC2626; border: 1px solid #FECACA; border-radius: 6px; padding: 5px 12px; font-weight: 600; } QPushButton:hover { background: #FEF2F2; border-color: #EF4444; }")
        self.clear_btn.clicked.connect(self.clear_all)
        self.clear_btn.hide()
        status_row.addWidget(self.clear_btn)
        layout.addLayout(status_row)
        
        # --- SPLIT VIEW ---
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setStyleSheet("QSplitter::handle { background: #E2E8F0; width: 1px; }")

        # LEFT: List
        left_cont = QWidget()
        left_layout = QVBoxLayout(left_cont)
        left_layout.setContentsMargins(0, 0, 15, 0)
        
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.SingleSelection)
        # Styled to be clean white with no outlines
        self.file_list.setStyleSheet("""
            QListWidget { border: 1px solid #CBD5E1; border-radius: 8px; outline: none; background: white; }
            QListWidget::item { border-bottom: 1px solid #F1F5F9; }
            QListWidget::item:selected { background-color: #EFF6FF; border: none; }
        """)
        self.file_list.currentRowChanged.connect(self.on_row_changed)
        left_layout.addWidget(self.file_list)
        
        # RIGHT: Preview
        right_cont = QWidget()
        right_cont.setStyleSheet("background: white; border: 1px solid #CBD5E1; border-radius: 8px;")
        right_layout = QVBoxLayout(right_cont)
        right_layout.setContentsMargins(20, 20, 20, 20)
        
        self.preview_lbl = ScalableImageLabel()
        self.preview_lbl.setStyleSheet("color: #94A3B8;")
        right_layout.addWidget(self.preview_lbl, 1)

        # Controls
        controls = QHBoxLayout()
        controls.setAlignment(Qt.AlignCenter)
        controls.setSpacing(25)
        
        def mk_icon_btn(name, tip, slot, destructive=False):
            b = QPushButton()
            c_hex, h_bg, h_bord = ("#DC2626", "#FEF2F2", "#EF4444") if destructive else ("#4B5563", "#F3F4F6", "#3B82F6")
            b.setIcon(create_painted_icon(name, c_hex, 64))
            b.setIconSize(QSize(28, 28))
            b.setFixedSize(50, 50)
            b.setCursor(Qt.PointingHandCursor)
            b.setToolTip(tip)
            b.setStyleSheet(f"QPushButton {{ background: white; border: 1px solid #E5E7EB; border-radius: 25px; }} QPushButton:hover {{ background: {h_bg}; border-color: {h_bord}; }}")
            b.clicked.connect(slot)
            return b

        self.btn_l = mk_icon_btn("rotate_left", "Rotate Left", lambda: self.rotate_curr(-90))
        self.btn_del = mk_icon_btn("trash", "Usuń Zdjęcie", self.delete_curr_selection, destructive=True)
        self.btn_r = mk_icon_btn("rotate_right", "Rotate Right", lambda: self.rotate_curr(90))
        
        for b in [self.btn_l, self.btn_del, self.btn_r]: b.setEnabled(False)
        controls.addWidget(self.btn_l); controls.addWidget(self.btn_del); controls.addWidget(self.btn_r)
        
        c_frame = QWidget(); c_frame.setLayout(controls); c_frame.setFixedHeight(80)
        right_layout.addWidget(c_frame, 0)

        self.splitter.addWidget(left_cont)
        self.splitter.addWidget(right_cont)
        # Set stretch factors: List 3/5, Preview Image 2/5
        self.splitter.setStretchFactor(0, 3)  # List gets 3 parts (60%)
        self.splitter.setStretchFactor(1, 2)  # Image gets 2 parts (40%)
        layout.addWidget(self.splitter, 1)

        # Transform Button
        self.tx_btn = QPushButton("✨ Przekształć zdjęcia na Excel")
        self.tx_btn.setFixedHeight(50)
        self.tx_btn.setCursor(Qt.PointingHandCursor)
        self.tx_btn.setEnabled(False)
        self.tx_btn.setStyleSheet("QPushButton { background: #2563EB; color: white; border-radius: 8px; font-weight: bold; font-size: 16px; } QPushButton:hover { background: #1D4ED8; } QPushButton:disabled { background: #E2E8F0; color: #94A3B8; }")
        self.tx_btn.clicked.connect(self.run_transform)
        layout.addWidget(self.tx_btn)

    # --- Logic ---
    def sel_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Wybierz Katalog")
        if d: self.add_from_dir(d)
    def sel_files(self):
        f, _ = QFileDialog.getOpenFileNames(self, "Wybierz Zdjęcia", "", "Zdjęcia (*.jpg *.png *.jpeg)")
        if f: self.handle_files(f)
    def handle_files(self, files):
        existing = set(self.selected_images)
        added = False
        for f in files:
            p = os.path.normpath(f)
            if p not in existing:
                self.selected_images.append(p)
                self.add_list_item(p)
                added = True
        if added:
            self.update_status()
            if self.file_list.currentRow() == -1: self.file_list.setCurrentRow(0)

    def add_from_dir(self, d):
        fs = []
        for e in ['*.jpg', '*.png', '*.jpeg']: fs.extend(glob.glob(os.path.join(d, e)))
        self.handle_files(fs)

    def add_list_item(self, path):
        # Create Item
        item = QListWidgetItem(self.file_list)
        item.setSizeHint(QSize(0, 56)) # Height for thumbnail + margins
        
        # Create Custom Widget (Thumb + Name + X)
        widget = FileListItemWidget(path, lambda: self.delete_file_by_path(path))
        self.file_list.setItemWidget(item, widget)

    def on_row_changed(self, row):
        if row < 0 or row >= len(self.selected_images):
            self.current_preview_index = -1
            self.preview_lbl.set_image(None)
            for b in [self.btn_l, self.btn_r, self.btn_del]: b.setEnabled(False)
            return
        self.current_preview_index = row
        for b in [self.btn_l, self.btn_r, self.btn_del]: b.setEnabled(True)
        self.refresh_preview()

    def refresh_preview(self):
        if self.current_preview_index == -1: return
        p = self.selected_images[self.current_preview_index]
        angle = self.image_rotations.get(p, 0)
        pix = QPixmap(p)
        if not pix.isNull() and angle != 0: pix = pix.transformed(QTransform().rotate(angle), Qt.SmoothTransformation)
        self.preview_lbl.set_image(pix)

    def rotate_curr(self, a):
        if self.current_preview_index == -1: return
        p = self.selected_images[self.current_preview_index]
        self.image_rotations[p] = (self.image_rotations.get(p, 0) + a) % 360
        self.refresh_preview()

    def delete_curr_selection(self):
        """Called by the Big Red Trash Button"""
        if self.current_preview_index == -1: return
        self.delete_at_index(self.current_preview_index)

    def delete_file_by_path(self, path):
        """Called by the small 'x' button in the list"""
        try:
            index = self.selected_images.index(path)
            self.delete_at_index(index)
        except ValueError:
            pass 

    def delete_at_index(self, index):
        if index < 0 or index >= len(self.selected_images): return
        
        path = self.selected_images.pop(index)
        if path in self.image_rotations: del self.image_rotations[path]
        
        self.file_list.takeItem(index) 
        self.update_status()
        
        if self.file_list.count() > 0:
            new_idx = min(index, self.file_list.count() - 1)
            self.file_list.setCurrentRow(new_idx)
        else:
            self.on_row_changed(-1)

    def clear_all(self):
        self.selected_images = []
        self.image_rotations = {}
        self.file_list.clear()
        self.on_row_changed(-1)
        self.update_status()

    def update_status(self):
        c = len(self.selected_images)
        if c > 0:
            self.status_label.setText(f"✓ Załadowano {c} zdjęci(e/a)")
            self.status_label.setStyleSheet("color: #059669; font-weight: 600;")
            self.clear_btn.show()
            self.tx_btn.setEnabled(True); self.tx_btn.setText(f"✨ Przekształć {c} zdjęci{'a' if c!=1 else 'e'}")
        else:
            self.status_label.setText("Nie załadowano zdjęć"); self.status_label.setStyleSheet("color: #64748B;")
            self.clear_btn.hide()
            self.tx_btn.setEnabled(False); self.tx_btn.setText("✨ Przekształć zdjęcia na Excel")

    def run_transform(self):
        if not self.selected_images: return
        if not config.get_gemini_api_key():
            return QMessageBox.critical(self, "Błąd", "GEMINI_API_KEY nie został ustawiony")
        
        import time
        timestamp = int(time.time() * 1000)  # Unique timestamp for this batch
        
        fl, tmp, preprocessing_failures = [], [], []
        for idx, p in enumerate(self.selected_images):
            r = self.image_rotations.get(p, 0)
            try:
                # Start with the original image
                current_path = p
                
                # Step 1: Apply rotation if needed and save to temp file
                if r != 0:
                    img = Image.open(p)
                    # Rotate in correct direction (negative because UI rotation is opposite)
                    rotated_img = img.rotate(-r, expand=True)
                    # Create unique temp filename with timestamp and index
                    base_name = os.path.splitext(os.path.basename(p))[0]
                    ext = os.path.splitext(p)[1]
                    rotated_path = os.path.join(os.path.dirname(p), f"tmp_{timestamp}_{idx}_rotated_{base_name}{ext}")
                    rotated_img.save(rotated_path)
                    current_path = rotated_path
                    tmp.append(rotated_path)
                    print(f"Rotated {os.path.basename(p)} by {r}° → {os.path.basename(rotated_path)}")
                
                # Step 2: Apply whitespace trimming
                base_name = os.path.splitext(os.path.basename(current_path))[0]
                ext = os.path.splitext(current_path)[1]
                trimmed_path_target = os.path.join(os.path.dirname(current_path), f"tmp_{timestamp}_{idx}_trimmed_{base_name}{ext}")
                trimmed_path = trim_whitespace(current_path, out_path=trimmed_path_target)
                
                if trimmed_path != current_path:
                    # Trimming created a new file
                    tmp.append(trimmed_path)
                    current_path = trimmed_path
                    print(f"Trimmed whitespace → {os.path.basename(trimmed_path)}")
                
                # Add final processed path to list for AI processing
                fl.append(current_path)
            except Exception as e:
                error_msg = f"Preprocessing error: {str(e)}"
                print(f"Error processing {p}: {error_msg}")
                preprocessing_failures.append((p, error_msg))
                # Don't add to fl - skip AI processing for failed preprocessing
        
        total_images = len(self.selected_images)
        processing_count = len(fl)
        
        if preprocessing_failures:
            msg = f"Przetwarzanie {processing_count} zdjęć ({len(preprocessing_failures)} niepowodzenia wstępne)..."
        else:
            msg = f"Przetwarzanie {processing_count} zdjęć..."
        
        self.pd = QProgressDialog(msg, "Anuluj", 0, processing_count, self)
        self.pd.setWindowModality(Qt.NonModal)  # Allow user to navigate while processing
        self.pd.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)  # Keep on top so it's visible
        self.pd.setAutoClose(True)
        self.pd.setAutoReset(True)
        self.pd.show()
        
        self.worker = TransformWorker(fl, max_workers=3)
        self.worker.progress.connect(self.pd.setValue)
        self.worker.finished.connect(lambda r: self.done(r, tmp, preprocessing_failures))
        self.worker.error.connect(lambda e: self.handle_worker_error(e, tmp))
        self.worker.start()

    def handle_worker_error(self, error_msg, tmp):
        """Handle worker thread errors."""
        if hasattr(self, 'pd') and self.pd:
            self.pd.close()
        # Clean up temp files
        for t in tmp:
            try:
                if os.path.exists(t):
                    os.remove(t)
            except Exception:
                pass
        QMessageBox.critical(self, "Błąd Przetwarzania", f"Błąd: {error_msg}")
    
    def done(self, res, tmp, preprocessing_failures):
        # Ensure progress dialog is closed
        if hasattr(self, 'pd') and self.pd:
            self.pd.close()
            self.pd = None
        # Clean up temporary processing files (rotation and trimming temps)
        for t in tmp:
            try:
                if os.path.exists(t):
                    os.remove(t)
            except Exception as e:
                print(f"Could not delete temp file {t}: {e}")
        
        # Count successes and failures
        successes = [r for r in res if r[1]]
        ai_failures = [r for r in res if not r[1]]
        total_failures = len(ai_failures) + len(preprocessing_failures)
        
        # Build detailed message
        msg = f"📊 Przetworzono {len(self.selected_images)} zdjęć(e/cia)\n"
        msg += f"✅ Pomyślnie przekształcono: {len(successes)}\n"
        if total_failures > 0:
            msg += f"❌ Niepowodzenia: {total_failures}\n"
        
        if successes:
            msg += "\n📁 Wygenerowane pliki:\n"
            for img_path, _, output_path, highlighted in successes[:5]:
                msg += f"  • {os.path.basename(output_path)}"
                if highlighted > 0:
                    msg += f" ({highlighted} rozbieżności)"
                msg += "\n"
            if len(successes) > 5:
                msg += f"  ... i {len(successes) - 5} więcej\n"
        
        if preprocessing_failures:
            msg += f"\n⚠️ Niepowodzenia wstępne ({len(preprocessing_failures)}):\n"
            for img_path, error in preprocessing_failures[:3]:
                msg += f"  • {os.path.basename(img_path)}:\n    {error}\n"
            if len(preprocessing_failures) > 3:
                msg += f"  ... i {len(preprocessing_failures) - 3} więcej\n"
        
        if ai_failures:
            msg += f"\n❌ Niepowodzenia przetwarzania AI ({len(ai_failures)}):\n"
            for img_path, _, error, _ in ai_failures[:3]:
                msg += f"  • {os.path.basename(img_path)}:\n    {error}\n"
            if len(ai_failures) > 3:
                msg += f"  ... i {len(ai_failures) - 3} więcej\n"
        
        QMessageBox.information(self, "Przekształcanie Zakończone", msg)
        
        # Remove successfully transformed images from the list
        for img_path, _, _, _ in successes:
            try:
                index = self.selected_images.index(img_path)
                self.delete_at_index(index)
            except (ValueError, IndexError):
                pass
        
        self.transformation_complete.emit()