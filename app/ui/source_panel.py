from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional

import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QComboBox, QSpinBox, QLineEdit, QToolButton,
    QAbstractItemView, QTableWidgetItem, QFrame
)
from PySide6.QtCore import Qt

from app.services.excel_io import list_sheets, load_table, ExcelTable
from app.ui.dnd_tables import SourceTable


@dataclass
class SourceState:
    src_id: str
    title: str
    table: Optional[ExcelTable] = None


class SourcePanel(QWidget):
    """
    Eine Quelle:
    - Datei / Sheet / Header / Match-Spalte (pro Quelle)
    - Suchleiste (exakte Zeichenfolge)
      - Nächstes: springt zum nächsten Treffer in den aktuell sichtbaren Zeilen
      - Filter: filtert die komplette Quelle nach Treffer (enthält)
      - Match-Zeilen: zurück zur Match-Auswahl (die zuletzt gerenderte Match-Menge)
    - 2 Tabellen (Top/Bottom) wenn >20 Spalten
    """

    def __init__(self, state: SourceState, parent=None):
        super().__init__(parent)
        self.state = state

        self._cols_top: List[str] = []
        self._cols_bottom: List[str] = []
        self._last_focus = None  # (view, row, col)
        self._match_df_cache = pd.DataFrame()

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # --- File line ---
        line = QHBoxLayout()
        self.btn_file = QPushButton(f"{state.title}: Datei wählen")
        self.lbl_path = QLabel("—")
        self.lbl_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.cb_sheet = QComboBox()
        self.sp_header = QSpinBox()
        self.sp_header.setMinimum(1)
        self.sp_header.setValue(1)

        # ✅ Match pro Quelle (Dropdown)
        self.cb_match = QComboBox()

        line.addWidget(self.btn_file)
        line.addWidget(self.lbl_path, 2)
        line.addWidget(QLabel("Sheet"))
        line.addWidget(self.cb_sheet, 1)
        line.addWidget(QLabel("Header"))
        line.addWidget(self.sp_header)
        line.addWidget(QLabel("Match"))
        line.addWidget(self.cb_match, 1)
        root.addLayout(line)

        # --- Search line ---
        sline = QHBoxLayout()
        self.ed_search = QLineEdit()
        self.ed_search.setPlaceholderText("Suche (exakte Zeichenfolge) …")

        self.btn_next = QToolButton()
        self.btn_next.setText("Nächstes")

        self.btn_clear = QToolButton()
        self.btn_clear.setText("Zurück")

        # ✅ in ganzer Quelle filtern
        self.btn_filter = QToolButton()
        self.btn_filter.setText("Filter")

        # ✅ zurück zu Match-Zeilen
        self.btn_back_match = QToolButton()
        self.btn_back_match.setText("Match-Zeilen")

        sline.addWidget(QLabel("Suche:"))
        sline.addWidget(self.ed_search, 2)
        sline.addWidget(self.btn_next)
        sline.addWidget(self.btn_clear)
        sline.addWidget(self.btn_filter)
        sline.addWidget(self.btn_back_match)
        root.addLayout(sline)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        root.addWidget(sep)

        # --- Tables ---
        self.view_top = SourceTable(0, 0)
        self.view_bottom = SourceTable(0, 0)

        for v in (self.view_top, self.view_bottom):
            v.setEditTriggers(QAbstractItemView.NoEditTriggers)
            v.setDragEnabled(True)
            v.setDragDropMode(QAbstractItemView.DragOnly)
            v.setSelectionMode(QAbstractItemView.ExtendedSelection)
            v.setSelectionBehavior(QAbstractItemView.SelectItems)
            v.cellClicked.connect(lambda r, c, vv=v: self._remember_focus(vv, r, c))

        root.addWidget(self.view_top, 3)
        root.addWidget(self.view_bottom, 2)
        self.view_bottom.hide()

        # events
        self.btn_file.clicked.connect(self._pick_file)
        self.cb_sheet.currentTextChanged.connect(self._reload_if_possible)
        self.sp_header.valueChanged.connect(lambda _: self._reload_if_possible())

        self.btn_next.clicked.connect(self.search_next)
        self.btn_clear.clicked.connect(self.restore_focus)
        self.ed_search.returnPressed.connect(self.search_next)

        self.btn_filter.clicked.connect(self.filter_full_table_exact)
        self.btn_back_match.clicked.connect(self.restore_match_rows)

    # ---------- load ----------
    def _pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, f"{self.state.title}: Excel wählen", "", "Excel (*.xlsx *.xls)")
        if not path:
            return
        self.lbl_path.setText(path)
        self.cb_sheet.clear()
        self.cb_sheet.addItems(list_sheets(path))
        self._reload(path)

    def _reload_if_possible(self):
        if not self.lbl_path.text() or self.lbl_path.text() == "—":
            return
        self._reload(self.lbl_path.text())

    def _reload(self, path: str):
        sheet = self.cb_sheet.currentText()
        if not sheet:
            return
        header = int(self.sp_header.value())
        self.state.table = load_table(path, sheet, header)

        # Match-Dropdown füllen
        self.cb_match.blockSignals(True)
        self.cb_match.clear()
        if self.state.table is not None:
            cols = [c for c in self.state.table.df.columns if c != "_KEY_"]
            self.cb_match.addItems([""] + cols)
        self.cb_match.blockSignals(False)

    # ---------- render ----------
    def render_rows(self, df: pd.DataFrame):
        # df ist die aktuell "gematchte" Menge (wird für "Match-Zeilen" gemerkt)
        self._match_df_cache = df.copy()

        all_cols = [c for c in df.columns if c != "_KEY_"]
        self._cols_top = all_cols[:20]
        self._cols_bottom = all_cols[20:]

        def render(view, cols):
            view.setColumnCount(len(cols))
            view.setHorizontalHeaderLabels(cols)
            view.setRowCount(len(df))
            for r in range(len(df)):
                rr = df.iloc[r]
                for c, col in enumerate(cols):
                    view.setItem(r, c, QTableWidgetItem("" if rr.get(col) is None else str(rr.get(col))))

        render(self.view_top, self._cols_top)

        if len(all_cols) > 20:
            self.view_bottom.show()
            render(self.view_bottom, self._cols_bottom)
        else:
            self.view_bottom.hide()
            self.view_bottom.setRowCount(0)
            self.view_bottom.setColumnCount(0)

        self._last_focus = None

    # ---------- match rows restore ----------
    def restore_match_rows(self):
        self.render_rows(self._match_df_cache)

    # ---------- search (visible rows) ----------
    def _remember_focus(self, view, r: int, c: int):
        self._last_focus = (view, r, c)

    def restore_focus(self):
        if not self._last_focus:
            return
        view, r, c = self._last_focus
        view.setCurrentCell(r, c)
        it = view.item(r, c)
        if it:
            view.scrollToItem(it)

    def search_next(self):
        needle = self.ed_search.text()
        if not needle:
            return

        views = [self.view_top] + ([self.view_bottom] if self.view_bottom.isVisible() else [])
        start_view_idx = 0
        start_r, start_c = 0, -1

        cur = None
        for vi, v in enumerate(views):
            if v.currentRow() >= 0 and v.currentColumn() >= 0 and v.hasFocus():
                cur = (vi, v.currentRow(), v.currentColumn())
                break
        if cur:
            start_view_idx, start_r, start_c = cur

        # scan forward
        for vi in range(start_view_idx, len(views)):
            v = views[vi]
            r0 = start_r if vi == start_view_idx else 0
            for r in range(r0, v.rowCount()):
                c0 = start_c + 1 if (vi == start_view_idx and r == r0) else 0
                for c in range(c0, v.columnCount()):
                    it = v.item(r, c)
                    if it and needle in it.text():
                        v.setCurrentCell(r, c)
                        v.scrollToItem(it)
                        v.setFocus()
                        return

        # wrap
        for vi in range(0, start_view_idx + 1):
            v = views[vi]
            for r in range(0, v.rowCount()):
                for c in range(0, v.columnCount()):
                    it = v.item(r, c)
                    if it and needle in it.text():
                        v.setCurrentCell(r, c)
                        v.scrollToItem(it)
                        v.setFocus()
                        return

    # ---------- filter full source table ----------
    def filter_full_table_exact(self):
        needle = self.ed_search.text()
        if not needle:
            return
        if self.state.table is None:
            return

        df = self.state.table.df

        # contains (exakt substring) in irgendeiner Zelle
        tmp = df.astype(str)
        mask = tmp.apply(lambda col: col.str.contains(needle, na=False), axis=0).any(axis=1)
        hits = df.loc[mask].copy()

        self.render_rows(hits.reset_index(drop=True))
