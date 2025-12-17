from __future__ import annotations
from dataclasses import dataclass
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
    QHBoxLayout, QPushButton, QComboBox, QCheckBox, QLineEdit, QMessageBox
)
from PySide6.QtCore import Qt

@dataclass
class ProposedChange:
    key: str
    t1_row_index: int
    target_col: str
    new_value: str
    source_info: str

class DetailDialog(QDialog):
    def __init__(self, key: str, t1_row_index: int, t1_row: dict, t2_df, t1_columns: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Detail – KEY {key}")
        self.key = key
        self.t1_row_index = t1_row_index
        self.t1_columns = t1_columns
        self.t1_row = t1_row
        self.t2_df = t2_df
        self.changes: list[ProposedChange] = []

        root = QVBoxLayout(self)

        root.addWidget(QLabel("Tabelle 1 – aktuelle Zeile"))
        self.t1_table = QTableWidget(1, len(t1_columns))
        self.t1_table.setHorizontalHeaderLabels(t1_columns)
        for c, col in enumerate(t1_columns):
            v = "" if t1_row.get(col) is None else str(t1_row.get(col))
            self.t1_table.setItem(0, c, QTableWidgetItem(v))
        self.t1_table.setEditTriggers(QTableWidget.NoEditTriggers)
        root.addWidget(self.t1_table)

        root.addWidget(QLabel("Tabelle 2 – passende Zeile(n) (pro Zelle auswählbar)"))
        self.t2_table = QTableWidget(0, 0)
        root.addWidget(self.t2_table)

        self.btn_apply = QPushButton("Auswahl in Änderungs-Liste übernehmen")
        self.btn_close = QPushButton("Schließen")
        btns = QHBoxLayout()
        btns.addWidget(self.btn_apply)
        btns.addWidget(self.btn_close)
        root.addLayout(btns)

        self.btn_close.clicked.connect(self.reject)
        self.btn_apply.clicked.connect(self.collect_changes)

        self.build_t2_table()

    def build_t2_table(self):
        # Spalten: pro T2-Spalte ein Block: [Use?] [T2 value] [Target(T1)] [Mode] [ManualValue]
        t2_cols = [c for c in self.t2_df.columns if c != "_KEY_"]

        headers = []
        self.col_meta = []  # (t2_col, kind)
        # Layout je T2-Spalte: Use | Value | Target | Mode | Manual
        for t2c in t2_cols:
            headers += [f"Use {t2c}", f"Value {t2c}", f"→ T1-Spalte", "Mode", "Manuell"]
            self.col_meta += [(t2c, "use"), (t2c, "value"), (t2c, "target"), (t2c, "mode"), (t2c, "manual")]

        self.t2_table.setColumnCount(len(headers))
        self.t2_table.setHorizontalHeaderLabels(headers)
        self.t2_table.setRowCount(len(self.t2_df))

        for r in range(len(self.t2_df)):
            row = self.t2_df.iloc[r]
            for c, (t2c, kind) in enumerate(self.col_meta):
                if kind == "use":
                    cb = QCheckBox()
                    cb.setChecked(False)
                    cb.setStyleSheet("margin-left:6px;")
                    self.t2_table.setCellWidget(r, c, cb)

                elif kind == "value":
                    v = "" if row.get(t2c) is None else str(row.get(t2c))
                    item = QTableWidgetItem(v)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.t2_table.setItem(r, c, item)

                elif kind == "target":
                    dd = QComboBox()
                    dd.addItem("")  # leer = noch kein Ziel
                    dd.addItems(self.t1_columns)
                    self.t2_table.setCellWidget(r, c, dd)

                elif kind == "mode":
                    dd = QComboBox()
                    dd.addItems(["1:1", "manuell"])
                    self.t2_table.setCellWidget(r, c, dd)

                elif kind == "manual":
                    le = QLineEdit()
                    le.setPlaceholderText("nur bei 'manuell' relevant")
                    self.t2_table.setCellWidget(r, c, le)

        self.t2_table.resizeColumnsToContents()

    def collect_changes(self):
        changes: list[ProposedChange] = []
        t2_cols = [c for c in self.t2_df.columns if c != "_KEY_"]

        for r in range(self.t2_table.rowCount()):
            # jede T2-Spalte prüfen
            for i, t2c in enumerate(t2_cols):
                base = i * 5
                cb: QCheckBox = self.t2_table.cellWidget(r, base)  # use
                if not cb.isChecked():
                    continue

                value_item = self.t2_table.item(r, base + 1)
                source_value = "" if value_item is None else value_item.text()

                target_dd: QComboBox = self.t2_table.cellWidget(r, base + 2)
                target_col = target_dd.currentText().strip()
                if not target_col:
                    QMessageBox.warning(self, "Fehlt", f"Zielspalte fehlt für T2-Spalte '{t2c}' in Zeile {r+1}.")
                    return

                mode_dd: QComboBox = self.t2_table.cellWidget(r, base + 3)
                mode = mode_dd.currentText()

                manual_le: QLineEdit = self.t2_table.cellWidget(r, base + 4)
                manual_value = manual_le.text().strip()

                new_value = source_value if mode == "1:1" else manual_value
                if mode == "manuell" and new_value == "":
                    QMessageBox.warning(self, "Fehlt", f"Manueller Wert ist leer für Ziel '{target_col}' (Zeile {r+1}).")
                    return

                changes.append(ProposedChange(
                    key=self.key,
                    t1_row_index=self.t1_row_index,
                    target_col=target_col,
                    new_value=new_value,
                    source_info=f"T2[{r}]:{t2c} ({mode})"
                ))

        self.changes = changes
        QMessageBox.information(self, "OK", f"{len(changes)} Änderungen gesammelt. Du kannst jetzt im Hauptfenster übernehmen.")
        self.accept()
