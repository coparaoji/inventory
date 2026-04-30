from __future__ import annotations

from pathlib import Path

from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QLineEdit,
    QLabel,
    QStatusBar,
    QHeaderView,
    QFileDialog,
    QMessageBox,
    QAbstractItemView,
    QSplitter,
    QComboBox,
    QSpinBox,
    QSizePolicy,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QResizeEvent

from app.docx_export import build_docx, QueueEntry, PRODUCT_TYPES

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff", ".tif"}


# ---------------------------------------------------------------------------
# Custom preview label (rescales pixmap on resize)
# ---------------------------------------------------------------------------

class _PreviewLabel(QLabel):
    """QLabel that keeps its pixmap scaled to fill the available area."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._source: QPixmap | None = None

    def set_source(self, pixmap: QPixmap | None) -> None:
        self._source = pixmap
        if pixmap:
            self._rescale()
        else:
            super().setPixmap(QPixmap())

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        if self._source:
            self._rescale()

    def _rescale(self) -> None:
        if self._source and not self._source.isNull():
            scaled = self._source.scaled(
                self.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            super().setPixmap(scaled)


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Inventory")
        self.setMinimumSize(1080, 720)
        self._current_folder: Path | None = None
        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        main_split = QSplitter(Qt.Orientation.Vertical)
        root.addWidget(main_split, stretch=1)

        # ── top half: folder browser ──────────────────────────────────────
        browser = QWidget()
        b_layout = QVBoxLayout(browser)
        b_layout.setContentsMargins(0, 0, 0, 0)
        b_layout.setSpacing(6)

        # Folder row
        folder_row = QHBoxLayout()
        self.folder_label = QLabel("No folder selected")
        self.folder_label.setStyleSheet("color: grey;")
        folder_btn = QPushButton("Select Folder…")
        folder_btn.clicked.connect(self._on_select_folder)
        folder_row.addWidget(folder_btn)
        folder_row.addWidget(self.folder_label, stretch=1)
        b_layout.addLayout(folder_row)

        # Search + add row
        search_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter by filename…")
        self.search_input.textChanged.connect(self._on_search)

        self.bulk_type_combo = QComboBox()
        self.bulk_type_combo.addItems(PRODUCT_TYPES)
        self.bulk_type_combo.setToolTip("Product type applied when adding images to queue")

        add_btn = QPushButton("Add Selected to Queue →")
        add_btn.setToolTip("Add highlighted rows to the print queue (multi-select with Ctrl / Shift)")
        add_btn.clicked.connect(self._on_add_to_queue)
        search_row.addWidget(QLabel("Search:"))
        search_row.addWidget(self.search_input, stretch=1)
        search_row.addWidget(QLabel("Type:"))
        search_row.addWidget(self.bulk_type_combo)
        search_row.addWidget(add_btn)
        b_layout.addLayout(search_row)

        # Horizontal split: file table | image preview
        h_split = QSplitter(Qt.Orientation.Horizontal)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Filename", "Extension", "Size"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSortingEnabled(True)
        self.table.selectionModel().selectionChanged.connect(self._on_table_selection_changed)
        h_split.addWidget(self.table)

        # Preview panel
        preview_panel = QWidget()
        preview_panel.setMinimumWidth(180)
        p_layout = QVBoxLayout(preview_panel)
        p_layout.setContentsMargins(6, 0, 0, 0)
        p_layout.setSpacing(4)

        self.preview_label = _PreviewLabel()
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setWordWrap(True)
        self.preview_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.preview_name = QLabel("")
        self.preview_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_name.setWordWrap(True)
        self.preview_name.setStyleSheet("font-size: 11px; color: #bbb;")

        p_layout.addWidget(self.preview_label, stretch=1)
        p_layout.addWidget(self.preview_name)
        h_split.addWidget(preview_panel)
        self._set_preview_idle()

        h_split.setSizes([660, 220])
        b_layout.addWidget(h_split, stretch=1)
        main_split.addWidget(browser)

        # ── bottom half: print queue ──────────────────────────────────────
        queue_panel = QWidget()
        q_layout = QVBoxLayout(queue_panel)
        q_layout.setContentsMargins(0, 4, 0, 0)
        q_layout.setSpacing(6)

        # Queue header
        q_header = QHBoxLayout()
        q_title = QLabel("Print Queue")
        q_title.setStyleSheet("font-weight: bold; font-size: 13px;")
        self.queue_count_label = QLabel("0 image(s)")
        self.queue_count_label.setStyleSheet("color: grey;")

        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._on_remove_from_queue)
        clear_btn = QPushButton("Clear Queue")
        clear_btn.clicked.connect(self._on_clear_queue)

        self.print_btn = QPushButton("Print to DOCX…")
        self.print_btn.setEnabled(False)
        self.print_btn.setStyleSheet(
            "QPushButton         { background-color: #2d7d46; color: white;"
            "                      font-weight: bold; padding: 4px 14px; }"
            "QPushButton:disabled{ background-color: #555; color: #888; }"
        )
        self.print_btn.clicked.connect(self._on_print_docx)

        q_header.addWidget(q_title)
        q_header.addWidget(self.queue_count_label)
        q_header.addStretch()
        q_header.addWidget(remove_btn)
        q_header.addWidget(clear_btn)
        q_header.addWidget(self.print_btn)
        q_layout.addLayout(q_header)

        # Queue table  (Filename | Product Type | Qty | hidden-path)
        self.queue_table = QTableWidget(0, 4)
        self.queue_table.setHorizontalHeaderLabels(
            ["Filename", "Product Type", "Qty", "_path"]
        )
        self.queue_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.queue_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        self.queue_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.ResizeToContents
        )
        self.queue_table.setColumnHidden(3, True)
        self.queue_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.queue_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.queue_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.queue_table.setMaximumHeight(200)
        q_layout.addWidget(self.queue_table)

        main_split.addWidget(queue_panel)
        main_split.setSizes([470, 260])

        # Status bar
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self._refresh_status()

    # ------------------------------------------------------------------
    # Slots — folder browser
    # ------------------------------------------------------------------

    def _on_select_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Select a folder", str(Path.home())
        )
        if not folder:
            return
        self._current_folder = Path(folder)
        self.folder_label.setText(folder)
        self.folder_label.setStyleSheet("")
        self._scan_folder()

    def _scan_folder(self) -> None:
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        self._set_preview_idle()

        if not self._current_folder:
            return

        images = [
            p for p in self._current_folder.rglob("*")
            if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS
        ]

        for img in images:
            row = self.table.rowCount()
            self.table.insertRow(row)

            name_item = QTableWidgetItem(img.name)
            name_item.setToolTip(str(img))
            name_item.setData(Qt.ItemDataRole.UserRole, str(img))

            ext_item = QTableWidgetItem(img.suffix.lower())

            size_bytes = img.stat().st_size
            size_item = QTableWidgetItem(_human_size(size_bytes))
            size_item.setData(Qt.ItemDataRole.UserRole, size_bytes)

            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, ext_item)
            self.table.setItem(row, 2, size_item)

        self.table.setSortingEnabled(True)
        self._refresh_status()

    def _on_search(self, text: str) -> None:
        text = text.lower()
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            match = text in (item.text().lower() if item else "")
            self.table.setRowHidden(row, not match if text else False)

    def _on_table_selection_changed(self) -> None:
        selected = self.table.selectionModel().selectedRows()
        if len(selected) == 1:
            item = self.table.item(selected[0].row(), 0)
            if item:
                self._show_preview(Path(item.data(Qt.ItemDataRole.UserRole)))
                return
        self._set_preview_idle()

    def _show_preview(self, path: Path) -> None:
        px = QPixmap(str(path))
        if px.isNull():
            self._set_preview_idle("Could not load image.")
            return
        self.preview_label.setStyleSheet("border: 1px solid #555; border-radius: 4px;")
        self.preview_label.setText("")
        self.preview_label.set_source(px)
        self.preview_name.setText(path.name)

    def _set_preview_idle(self, msg: str = "Select one image to preview") -> None:
        self.preview_label.set_source(None)
        self.preview_label.setText(msg)
        self.preview_label.setStyleSheet(
            "color: grey; border: 1px solid #444; border-radius: 4px;"
        )
        self.preview_name.setText("")

    # ------------------------------------------------------------------
    # Slots — print queue
    # ------------------------------------------------------------------

    def _on_add_to_queue(self) -> None:
        selected = self.table.selectionModel().selectedRows()
        if not selected:
            self.status.showMessage(
                "No images selected — highlight rows first (Ctrl/Shift for multi-select).", 3000
            )
            return

        existing = self._queue_paths()
        added = 0

        for index in selected:
            item = self.table.item(index.row(), 0)
            if not item:
                continue
            path = Path(item.data(Qt.ItemDataRole.UserRole))
            if path in existing:
                continue

            row = self.queue_table.rowCount()
            self.queue_table.insertRow(row)

            # Filename (read-only)
            fname = QTableWidgetItem(path.name)
            fname.setToolTip(str(path))
            self.queue_table.setItem(row, 0, fname)

            # Product type combobox — initialised to the bulk selector choice
            combo = QComboBox()
            combo.addItems(PRODUCT_TYPES)
            combo.setCurrentText(self.bulk_type_combo.currentText())
            self.queue_table.setCellWidget(row, 1, combo)

            # Quantity spinbox
            spin = QSpinBox()
            spin.setMinimum(1)
            spin.setMaximum(999)
            spin.setValue(1)
            self.queue_table.setCellWidget(row, 2, spin)

            # Hidden full path
            self.queue_table.setItem(row, 3, QTableWidgetItem(str(path)))

            added += 1

        self._refresh_queue_ui()
        total = self.queue_table.rowCount()
        self.status.showMessage(
            f"Added {added} image(s) to queue.  Queue total: {total}.", 3000
        )

    def _on_remove_from_queue(self) -> None:
        rows = sorted(
            {idx.row() for idx in self.queue_table.selectionModel().selectedRows()},
            reverse=True,
        )
        for row in rows:
            self.queue_table.removeRow(row)
        self._refresh_queue_ui()

    def _on_clear_queue(self) -> None:
        self.queue_table.setRowCount(0)
        self._refresh_queue_ui()

    def _on_print_docx(self) -> None:
        entries = self._collect_entries()
        if not entries:
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save image grid",
            str(Path.cwd() / "image_grid.docx"),
            "Word Document (*.docx)",
        )
        if not save_path:
            return

        try:
            build_docx(entries, Path(save_path))
            total_slots = sum(e.quantity for e in entries)
            QMessageBox.information(
                self,
                "Saved",
                f"Exported {total_slots} image slot(s) across "
                f"{len(entries)} queue item(s).\n\n{save_path}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export failed", str(exc))

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _queue_paths(self) -> set[Path]:
        paths: set[Path] = set()
        for row in range(self.queue_table.rowCount()):
            item = self.queue_table.item(row, 3)
            if item:
                paths.add(Path(item.text()))
        return paths

    def _collect_entries(self) -> list[QueueEntry]:
        entries: list[QueueEntry] = []
        for row in range(self.queue_table.rowCount()):
            path_item = self.queue_table.item(row, 3)
            combo: QComboBox | None = self.queue_table.cellWidget(row, 1)  # type: ignore[assignment]
            spin: QSpinBox | None = self.queue_table.cellWidget(row, 2)    # type: ignore[assignment]
            if path_item and combo and spin:
                entries.append(
                    QueueEntry(
                        path=Path(path_item.text()),
                        product_type=combo.currentText(),
                        quantity=spin.value(),
                    )
                )
        return entries

    def _refresh_status(self) -> None:
        folder = str(self._current_folder) if self._current_folder else "none"
        self.status.showMessage(
            f"{self.table.rowCount()} image(s) found  |  folder: {folder}"
        )

    def _refresh_queue_ui(self) -> None:
        count = self.queue_table.rowCount()
        self.queue_count_label.setText(f"{count} image(s)")
        self.print_btn.setEnabled(count > 0)


# ---------------------------------------------------------------------------
# Misc helpers
# ---------------------------------------------------------------------------

def _human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"
