import os
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
)
from PyQt6.QtCore import Qt

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff", ".tif", ".svg"}


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Inventory")
        self.setMinimumSize(900, 600)
        self._current_folder: Path | None = None
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # --- folder row ---
        folder_row = QHBoxLayout()
        self.folder_label = QLabel("No folder selected")
        self.folder_label.setStyleSheet("color: grey;")
        folder_btn = QPushButton("Select Folder…")
        folder_btn.clicked.connect(self._on_select_folder)

        folder_row.addWidget(folder_btn)
        folder_row.addWidget(self.folder_label, stretch=1)
        root.addLayout(folder_row)

        # --- search row ---
        toolbar = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter by filename…")
        self.search_input.textChanged.connect(self._on_search)

        toolbar.addWidget(QLabel("Search:"))
        toolbar.addWidget(self.search_input, stretch=1)
        root.addLayout(toolbar)

        # --- table ---
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Filename", "Extension", "Size"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSortingEnabled(True)
        root.addWidget(self.table, stretch=1)

        # --- status bar ---
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self._refresh_status()

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------

    def _on_select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select a folder", str(Path.home()))
        if not folder:
            return
        self._current_folder = Path(folder)
        self.folder_label.setText(folder)
        self.folder_label.setStyleSheet("")
        self._scan_folder()

    def _scan_folder(self):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)

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

            ext_item = QTableWidgetItem(img.suffix.lower())

            size_bytes = img.stat().st_size
            size_item = QTableWidgetItem(_human_size(size_bytes))
            size_item.setData(Qt.ItemDataRole.UserRole, size_bytes)  # for correct numeric sort

            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, ext_item)
            self.table.setItem(row, 2, size_item)

        self.table.setSortingEnabled(True)
        self._refresh_status()

    def _on_search(self, text: str):
        text = text.lower()
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            visible = text in (item.text().lower() if item else "")
            self.table.setRowHidden(row, not visible if text else False)

    def _refresh_status(self):
        folder_str = str(self._current_folder) if self._current_folder else "none"
        self.status.showMessage(f"{self.table.rowCount()} image(s) found  |  folder: {folder_str}")


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"
