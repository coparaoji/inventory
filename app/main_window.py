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
)
from PyQt6.QtCore import Qt


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Inventory")
        self.setMinimumSize(900, 600)

        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # --- toolbar row ---
        toolbar = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search…")
        self.search_input.textChanged.connect(self._on_search)

        add_btn = QPushButton("+ Add")
        add_btn.clicked.connect(self._on_add)
        delete_btn = QPushButton("Delete")
        delete_btn.clicked.connect(self._on_delete)

        toolbar.addWidget(QLabel("Search:"))
        toolbar.addWidget(self.search_input, stretch=1)
        toolbar.addSpacing(12)
        toolbar.addWidget(add_btn)
        toolbar.addWidget(delete_btn)
        root.addLayout(toolbar)

        # --- table ---
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Name", "Quantity", "Notes"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked)
        root.addWidget(self.table, stretch=1)

        # --- status bar ---
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self._refresh_status()

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------

    def _on_add(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem("New item"))
        self.table.setItem(row, 1, QTableWidgetItem("0"))
        self.table.setItem(row, 2, QTableWidgetItem(""))
        self.table.editItem(self.table.item(row, 0))
        self._refresh_status()

    def _on_delete(self):
        selected = self.table.selectedItems()
        if not selected:
            return
        rows = sorted({item.row() for item in selected}, reverse=True)
        for row in rows:
            self.table.removeRow(row)
        self._refresh_status()

    def _on_search(self, text: str):
        text = text.lower()
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            match = text in (item.text().lower() if item else "")
            self.table.setRowHidden(row, not match if text else False)

    def _refresh_status(self):
        self.status.showMessage(f"{self.table.rowCount()} item(s)")
