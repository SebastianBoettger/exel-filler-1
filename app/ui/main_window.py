from __future__ import annotations

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QComboBox, QSpinBox, QTableWidgetItem,
    QMessageBox, QAbstractItemView,
    QInputDialog, QDialog, QFormLayout, QCheckBox, QDockWidget, QMenu
)
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QColor, QBrush, QAction

from app.services.excel_io import list_sheets, load_table
from app.services.matcher import MatchEngine
from app.services.normalize import is_missing
from app.services.transforms import split_street_house, normalize_phone, state_from_zip_de
from app.services.settings import AppSettings, load_settings, save_settings
from app.ui.dnd_tables import SourceTable, TargetTable


def _pick_text_color_for_bg(hex_color: str) -> QColor:
    c = QColor(hex_color)
    y = 0.2126 * c.red() + 0.7152 * c.green() + 0.0722 * c.blue()
    return QColor("#000000") if y > 140 else QColor("#FFFFFF")


def _palette_50() -> list[str]:
    return [
        "#e6194b", "#3cb44b", "#ffe119", "#4363d8", "#f58231",
        "#911eb4", "#46f0f0", "#f032e6", "#bcf60c", "#fabebe",
        "#008080", "#e6beff", "#9a6324", "#fffac8", "#800000",
        "#aaffc3", "#808000", "#ffd8b1", "#000075", "#808080",
        "#ffffff", "#000000", "#ff7f00", "#1f77b4", "#2ca02c",
        "#d62728", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f",
        "#bcbd22", "#17becf", "#4daf4a", "#984ea3", "#ff33cc",
        "#a65628", "#f781bf", "#999999", "#66c2a5", "#fc8d62",
        "#8da0cb", "#e78ac3", "#a6d854", "#ffd92f", "#e5c494",
        "#b3b3b3", "#006400", "#8b0000", "#00008b", "#ff1493",
    ]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel-Filler GUI")

        self.t1 = None
        self.t2 = None
        self.engine: MatchEngine | None = None

        self.keys_queue: list[str] = []
        self.current_pos = -1
        self.current_key: str | None = None

        self._t1_sheet_slot = None
        self._t2_sheet_slot = None

        # Settings
        self.settings: AppSettings = load_settings()
        self.col_links = dict(self.settings.col_links)
        self.cuts = dict(self.settings.cuts)
        self.country_default_value = self.settings.country_default_value

        # For 2-line split
        self._t1_cols_top: list[str] = []
        self._t1_cols_bottom: list[str] = []
        self._t2_cols_top: list[str] = []
        self._t2_cols_bottom: list[str] = []

        central = QWidget(self)
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # --------- File selectors ----------
        row1 = QHBoxLayout()
        self.btn_t1 = QPushButton("Datei 1 wählen")
        self.lbl_t1 = QLabel("—")
        self.cb_t1_sheet = QComboBox()
        self.sp_t1_header = QSpinBox()
        self.sp_t1_header.setMinimum(1)
        self.sp_t1_header.setValue(1)
        self.cb_t1_key = QComboBox()

        row1.addWidget(self.btn_t1)
        row1.addWidget(self.lbl_t1)
        row1.addWidget(QLabel("Sheet"))
        row1.addWidget(self.cb_t1_sheet)
        row1.addWidget(QLabel("Header"))
        row1.addWidget(self.sp_t1_header)
        row1.addWidget(QLabel("KEY"))
        row1.addWidget(self.cb_t1_key)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.btn_t2 = QPushButton("Datei 2 wählen")
        self.lbl_t2 = QLabel("—")
        self.cb_t2_sheet = QComboBox()
        self.sp_t2_header = QSpinBox()
        self.sp_t2_header.setMinimum(1)
        self.sp_t2_header.setValue(1)
        self.cb_t2_key = QComboBox()

        row2.addWidget(self.btn_t2)
        row2.addWidget(self.lbl_t2)
        row2.addWidget(QLabel("Sheet"))
        row2.addWidget(self.cb_t2_sheet)
        row2.addWidget(QLabel("Header"))
        row2.addWidget(self.sp_t2_header)
        row2.addWidget(QLabel("KEY"))
        row2.addWidget(self.cb_t2_key)
        layout.addLayout(row2)

        self.btn_start = QPushButton("Start (fehlende Kundennummern)")
        layout.addWidget(self.btn_start)

        # --------- Navigation / actions ----------
        nav = QHBoxLayout()
        self.btn_add_col = QPushButton("Spalte hinzufügen (T1)")
        self.btn_links = QPushButton("Kopplungen")
        self.btn_cuts = QPushButton("Cuts")
        self.btn_fill_row = QPushButton("Plausibel füllen (Zeile)")
        self.btn_fill_all = QPushButton("Plausibel füllen (Gesamt)")
        self.btn_prev = QPushButton("◀ Zurück")
        self.btn_next = QPushButton("Nächste ▶")
        self.btn_save_as = QPushButton("Speichern unter…")
        self.btn_save_inplace = QPushButton("Speichern (gleiche Datei)")
        self.btn_save = QPushButton("Speichern (neu)")

        nav.addWidget(self.btn_add_col)
        nav.addWidget(self.btn_links)
        nav.addWidget(self.btn_cuts)
        nav.addWidget(self.btn_fill_row)
        nav.addWidget(self.btn_fill_all)
        nav.addWidget(self.btn_prev)
        nav.addWidget(self.btn_next)
        nav.addWidget(self.btn_save_as)
        nav.addWidget(self.btn_save_inplace)
        nav.addWidget(self.btn_save)
        layout.addLayout(nav)

        # --------- T1 container: two rows ----------
        self.t1_container = QWidget()
        t1_layout = QVBoxLayout(self.t1_container)
        t1_layout.setContentsMargins(0, 0, 0, 0)

        self.t1_view_top = TargetTable(1, 0)
        self.t1_view_bottom = TargetTable(1, 0)

        for v in (self.t1_view_top, self.t1_view_bottom):
            v.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
            v.setDragDropMode(QAbstractItemView.DropOnly)
            v.horizontalHeader().setSectionsMovable(True)
            v.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
            v.horizontalHeader().customContextMenuRequested.connect(lambda pos, view=v: self._header_menu(view, "t1", pos))
            v.horizontalHeader().sectionMoved.connect(lambda *_: self._persist_order_from_views("t1"))
            v.itemChanged.connect(self.on_t1_item_changed_any)

        t1_layout.addWidget(self.t1_view_top)
        t1_layout.addWidget(self.t1_view_bottom)
        layout.addWidget(self.t1_container)

        # --------- Status ----------
        self.status = QLabel(f"Settings geladen (Country default: {self.country_default_value})")
        layout.addWidget(self.status)

        # --------- T2 dock: two rows ----------
        self.t2_container = QWidget()
        t2_layout = QVBoxLayout(self.t2_container)
        t2_layout.setContentsMargins(0, 0, 0, 0)

        self.t2_view_top = SourceTable(0, 0)
        self.t2_view_bottom = SourceTable(0, 0)

        for v in (self.t2_view_top, self.t2_view_bottom):
            v.setEditTriggers(QAbstractItemView.NoEditTriggers)
            v.setDragEnabled(True)
            v.setDragDropMode(QAbstractItemView.DragOnly)
            v.setSelectionMode(QAbstractItemView.ExtendedSelection)
            v.setSelectionBehavior(QAbstractItemView.SelectItems)
            v.horizontalHeader().setSectionsMovable(True)
            v.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
            v.horizontalHeader().customContextMenuRequested.connect(lambda pos, view=v: self._header_menu(view, "t2", pos))
            v.horizontalHeader().sectionMoved.connect(lambda *_: self._persist_order_from_views("t2"))
            v.cellDoubleClicked.connect(lambda r, c, view=v: self.quick_copy_from_t2(view, r, c))

        t2_layout.addWidget(self.t2_view_top)
        t2_layout.addWidget(self.t2_view_bottom)

        self.dock_t2 = QDockWidget("Tabelle 2 (Quelle)", self)
        self.dock_t2.setWidget(self.t2_container)
        self.dock_t2.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea | Qt.BottomDockWidgetArea)
        self.dock_t2.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_t2)
        self.setDockOptions(QMainWindow.AllowTabbedDocks | QMainWindow.AllowNestedDocks)

        # --------- Events ----------
        self.btn_t1.clicked.connect(self.pick_t1)
        self.btn_t2.clicked.connect(self.pick_t2)
        self.btn_start.clicked.connect(self.start_scan)

        self.btn_prev.clicked.connect(self.prev_key)
        self.btn_next.clicked.connect(self.next_key)

        self.btn_save.clicked.connect(self.save_new_file)
        self.btn_save_as.clicked.connect(self.save_as)
        self.btn_save_inplace.clicked.connect(self.save_inplace)

        self.btn_add_col.clicked.connect(self.add_column_t1_global)
        self.btn_links.clicked.connect(self.open_links_dialog)
        self.btn_cuts.clicked.connect(self.open_cuts_dialog)
        self.btn_fill_row.clicked.connect(self.autofill_current_key_linked)
        self.btn_fill_all.clicked.connect(self.autofill_all_linked)

    # ---------------- Settings persistence ----------------
    def _save_settings_now(self):
        self.settings = AppSettings(
            col_links=dict(self.col_links),
            cuts=dict(self.cuts),
            country_default_value=str(self.country_default_value),
            t1_hidden=list(self.settings.t1_hidden),
            t2_hidden=list(self.settings.t2_hidden),
            t1_colors=dict(self.settings.t1_colors),
            t2_colors=dict(self.settings.t2_colors),
            t1_order=list(self.settings.t1_order),
            t2_order=list(self.settings.t2_order),
        )
        save_settings(self.settings)

    def closeEvent(self, event):
        try:
            self._save_settings_now()
        except Exception:
            pass
        super().closeEvent(event)

    # ---------------- Header menu / prefs ----------------
    def _header_menu(self, table, table_key: str, pos: QPoint):
        header = table.horizontalHeader()
        logical_index = header.logicalIndexAt(pos)
        if logical_index < 0:
            return
        col_item = table.horizontalHeaderItem(logical_index)
        if not col_item:
            return
        col_name = col_item.text()

        menu = QMenu(self)

        a_hide = menu.addAction(f"Spalte ausblenden: {col_name}")
        a_hide.triggered.connect(lambda: self._hide_column(table_key, col_name))

        hidden = self.settings.t1_hidden if table_key == "t1" else self.settings.t2_hidden
        if hidden:
            sub = menu.addMenu("Ausgeblendete Spalten einblenden")
            for h in hidden:
                a = sub.addAction(h)
                a.triggered.connect(lambda _, name=h: self._show_column(table_key, name))

        menu.addSeparator()

        subc = menu.addMenu(f"Spaltenkopf-Farbe setzen: {col_name}")
        for hexcol in _palette_50():
            a = subc.addAction(hexcol)
            a.triggered.connect(lambda _, hc=hexcol: self._set_col_color(table_key, col_name, hc))

        a_clear = menu.addAction("Farbe löschen")
        a_clear.triggered.connect(lambda: self._clear_col_color(table_key, col_name))

        menu.addSeparator()

        a_show_all = menu.addAction("Alle Spalten einblenden")
        a_show_all.triggered.connect(lambda: self._show_all(table_key))

        a_reset_order = menu.addAction("Spaltenreihenfolge zurücksetzen")
        a_reset_order.triggered.connect(lambda: self._reset_order(table_key))

        menu.exec(header.mapToGlobal(pos))

    def _hide_column(self, table_key: str, col_name: str):
        hidden = self.settings.t1_hidden if table_key == "t1" else self.settings.t2_hidden
        if col_name not in hidden:
            hidden.append(col_name)
        self._apply_table_prefs(table_key)
        self._save_settings_now()

    def _show_column(self, table_key: str, col_name: str):
        hidden = self.settings.t1_hidden if table_key == "t1" else self.settings.t2_hidden
        if col_name in hidden:
            hidden.remove(col_name)
        self._apply_table_prefs(table_key)
        self._save_settings_now()

    def _show_all(self, table_key: str):
        if table_key == "t1":
            self.settings.t1_hidden = []
        else:
            self.settings.t2_hidden = []
        self._apply_table_prefs(table_key)
        self._save_settings_now()

    def _set_col_color(self, table_key: str, col_name: str, hex_color: str):
        colors = self.settings.t1_colors if table_key == "t1" else self.settings.t2_colors
        colors[col_name] = hex_color
        self._apply_table_prefs(table_key)
        self._save_settings_now()

    def _clear_col_color(self, table_key: str, col_name: str):
        colors = self.settings.t1_colors if table_key == "t1" else self.settings.t2_colors
        if col_name in colors:
            del colors[col_name]
        self._apply_table_prefs(table_key)
        self._save_settings_now()

    def _persist_order_from_views(self, table_key: str):
        if table_key == "t1":
            order = self._get_visual_order(self.t1_view_top) + self._get_visual_order(self.t1_view_bottom)
            self.settings.t1_order = [c for c in order if c]
        else:
            order = self._get_visual_order(self.t2_view_top) + self._get_visual_order(self.t2_view_bottom)
            self.settings.t2_order = [c for c in order if c]
        self._save_settings_now()

    def _reset_order(self, table_key: str):
        if table_key == "t1":
            self.settings.t1_order = []
        else:
            self.settings.t2_order = []
        self._apply_table_prefs(table_key)
        self._save_settings_now()

    def _get_visual_order(self, table) -> list[str]:
        header = table.horizontalHeader()
        names = []
        for visual in range(header.count()):
            logical = header.logicalIndex(visual)
            it = table.horizontalHeaderItem(logical)
            if it:
                names.append(it.text())
        return names

    def _apply_table_prefs(self, table_key: str):
        if table_key == "t1":
            views = (self.t1_view_top, self.t1_view_bottom)
            hidden = self.settings.t1_hidden
            colors = self.settings.t1_colors
            order = self.settings.t1_order
        else:
            views = (self.t2_view_top, self.t2_view_bottom)
            hidden = self.settings.t2_hidden
            colors = self.settings.t2_colors
            order = self.settings.t2_order

        # Hide/show + colors
        for table in views:
            for i in range(table.columnCount()):
                it = table.horizontalHeaderItem(i)
                if not it:
                    continue
                name = it.text()
                table.setColumnHidden(i, name in hidden)
                if name in colors:
                    bg = QColor(colors[name])
                    fg = _pick_text_color_for_bg(colors[name])
                    it.setBackground(QBrush(bg))
                    it.setForeground(QBrush(fg))
                else:
                    it.setBackground(QBrush())
                    it.setForeground(QBrush())

        # Order (best-effort, only within each view)
        if order:
            for table in views:
                header = table.horizontalHeader()
                existing = self._get_visual_order(table)
                if not existing:
                    continue
                # subset order
                desired = [c for c in order if c in existing]
                for target_visual, name in enumerate(desired):
                    for v in range(header.count()):
                        logical = header.logicalIndex(v)
                        it = table.horizontalHeaderItem(logical)
                        if it and it.text() == name:
                            header.moveSection(v, target_visual)
                            break

    # ---------------- Load files ----------------
    def pick_t1(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel Datei 1", "", "Excel (*.xlsx)")
        if not path:
            return
        self.lbl_t1.setText(path)
        self.cb_t1_sheet.clear()
        self.cb_t1_sheet.addItems(list_sheets(path))

        if self._t1_sheet_slot is not None:
            try:
                self.cb_t1_sheet.currentTextChanged.disconnect(self._t1_sheet_slot)
            except Exception:
                pass
        self._t1_sheet_slot = lambda _=None, p=path: self.load_t1(p)
        self.cb_t1_sheet.currentTextChanged.connect(self._t1_sheet_slot)

        self.load_t1(path)

    def load_t1(self, path: str):
        self.t1 = load_table(path, self.cb_t1_sheet.currentText(), self.sp_t1_header.value())
        self.cb_t1_key.clear()
        self.cb_t1_key.addItems(self.t1.df.columns.tolist())

    def pick_t2(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel Datei 2", "", "Excel (*.xlsx)")
        if not path:
            return
        self.lbl_t2.setText(path)
        self.cb_t2_sheet.clear()
        self.cb_t2_sheet.addItems(list_sheets(path))

        if self._t2_sheet_slot is not None:
            try:
                self.cb_t2_sheet.currentTextChanged.disconnect(self._t2_sheet_slot)
            except Exception:
                pass
        self._t2_sheet_slot = lambda _=None, p=path: self.load_t2(p)
        self.cb_t2_sheet.currentTextChanged.connect(self._t2_sheet_slot)

        self.load_t2(path)

    def load_t2(self, path: str):
        self.t2 = load_table(path, self.cb_t2_sheet.currentText(), self.sp_t2_header.value())
        self.cb_t2_key.clear()
        self.cb_t2_key.addItems(self.t2.df.columns.tolist())

    # ---------------- Scan ----------------
    def start_scan(self):
        if not self.t1 or not self.t2:
            QMessageBox.warning(self, "Fehlt", "Bitte beide Tabellen laden.")
            return

        self.engine = MatchEngine(
            self.t1.df, self.cb_t1_key.currentText(),
            self.t2.df, self.cb_t2_key.currentText()
        )

        cols = [c for c in self.engine.df1.columns if c not in ["_KEY_", self.cb_t1_key.currentText()]]
        self.keys_queue = self.engine.keys_with_missing(cols)
        self.current_pos = -1

        self.status.setText(f"{len(self.keys_queue)} Kundennummern mit Lücken gefunden")
        self.next_key()

    # ---------------- Navigation ----------------
    def next_key(self):
        if not self.keys_queue:
            QMessageBox.information(self, "Fertig", "Keine fehlenden Felder gefunden.")
            return
        if self.current_pos < len(self.keys_queue) - 1:
            self.current_pos += 1
            self.show_key(self.keys_queue[self.current_pos])

    def prev_key(self):
        if not self.keys_queue:
            return
        if self.current_pos > 0:
            self.current_pos -= 1
            self.show_key(self.keys_queue[self.current_pos])

    # ---------------- Show key with 2-line split ----------------
    def show_key(self, key: str):
        self.current_key = key
        idx = self.engine.t1_row_index_for_key(key)

        all_t1_cols = [c for c in self.engine.df1.columns if c != "_KEY_"]
        self._t1_cols_top = all_t1_cols[:20]
        self._t1_cols_bottom = all_t1_cols[20:]

        row = self.engine.df1.loc[idx]

        def render_t1(view, cols):
            view.blockSignals(True)
            view.setColumnCount(len(cols))
            view.setHorizontalHeaderLabels(cols)
            view.setRowCount(1)
            for i, col in enumerate(cols):
                view.setItem(0, i, QTableWidgetItem("" if row.get(col) is None else str(row.get(col))))
            view.blockSignals(False)

        render_t1(self.t1_view_top, self._t1_cols_top)
        if len(all_t1_cols) > 20:
            self.t1_view_bottom.show()
            render_t1(self.t1_view_bottom, self._t1_cols_bottom)
        else:
            self.t1_view_bottom.hide()

        df2 = self.engine.t2_rows_for_key(key)
        all_t2_cols = [c for c in df2.columns if c != "_KEY_"]
        self._t2_cols_top = all_t2_cols[:20]
        self._t2_cols_bottom = all_t2_cols[20:]

        def render_t2(view, cols):
            view.setColumnCount(len(cols))
            view.setHorizontalHeaderLabels(cols)
            view.setRowCount(len(df2))
            for r in range(len(df2)):
                for c, col in enumerate(cols):
                    view.setItem(r, c, QTableWidgetItem("" if df2.iloc[r].get(col) is None else str(df2.iloc[r].get(col))))

        render_t2(self.t2_view_top, self._t2_cols_top)
        if len(all_t2_cols) > 20:
            self.t2_view_bottom.show()
            render_t2(self.t2_view_bottom, self._t2_cols_bottom)
        else:
            self.t2_view_bottom.hide()

        self.status.setText(f"KEY {key} ({self.current_pos+1}/{len(self.keys_queue)})")

        self._apply_table_prefs("t1")
        self._apply_table_prefs("t2")

        self.t1_view_top.setCurrentCell(0, 0)

    # ---------------- Sync edits / drops ----------------
    def on_t1_item_changed_any(self, item: QTableWidgetItem):
        if not self.engine or self.current_key is None:
            return
        t1_idx = self.engine.t1_row_index_for_key(self.current_key)
        if t1_idx is None:
            return

        view = item.tableWidget()
        if view == self.t1_view_top:
            col_name = self._t1_cols_top[item.column()]
        else:
            col_name = self._t1_cols_bottom[item.column()]

        self.engine.df1.at[t1_idx, col_name] = item.text()

    def on_t1_cell_dropped(self, view, row: int, col: int, text: str):
        if not self.engine or self.current_key is None:
            return
        t1_idx = self.engine.t1_row_index_for_key(self.current_key)
        if t1_idx is None:
            return

        if view == self.t1_view_top:
            col_name = self._t1_cols_top[col]
        else:
            col_name = self._t1_cols_bottom[col]

        self.engine.df1.at[t1_idx, col_name] = text

    def quick_copy_from_t2(self, t2_view, r: int, c: int):
        item = t2_view.item(r, c)
        if not item:
            return
        text = item.text()

        target_view = self.t1_view_top
        target_col = target_view.currentColumn()
        if target_col < 0:
            target_col = 0

        target_view.blockSignals(True)
        target_view.setItem(0, target_col, QTableWidgetItem(text))
        target_view.blockSignals(False)
        self.on_t1_cell_dropped(target_view, 0, target_col, text)

    # ---------------- Save buttons ----------------
    def save_new_file(self):
        if not self.engine or not self.t1:
            return
        from app.services.apply_changes import save_filled
        out = save_filled(self.engine.df1, self.t1.path.parent, self.t1.path.stem)
        QMessageBox.information(self, "Gespeichert", str(out))

    def save_as(self):
        if not self.engine:
            return
        from app.services.apply_changes import save_to_path
        default = "output.xlsx"
        if self.t1:
            default = str(self.t1.path.with_name(self.t1.path.stem + "_filled.xlsx"))
        path, _ = QFileDialog.getSaveFileName(self, "Speichern unter…", default, "Excel (*.xlsx)")
        if not path:
            return
        out = save_to_path(self.engine.df1, path)
        QMessageBox.information(self, "Gespeichert", str(out))

    def save_inplace(self):
        if not self.engine or not self.t1:
            return
        from app.services.apply_changes import save_in_place
        try:
            out = save_in_place(self.engine.df1, self.t1.path, self.t1.sheet, make_backup=True)
            QMessageBox.information(self, "Gespeichert", f"In Datei gespeichert (Backup erstellt):\n{out}")
        except PermissionError:
            QMessageBox.critical(self, "Fehler", "Datei ist vermutlich in Excel geöffnet. Bitte schließen und erneut speichern.")
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Speichern fehlgeschlagen:\n{e}")

    # ---------------- Copplings ----------------
    def open_links_dialog(self):
        if not self.engine:
            QMessageBox.warning(self, "Fehlt", "Bitte erst Start ausführen.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Spalten-Kopplungen (T1 → T2)")
        layout = QVBoxLayout(dlg)
        form = QFormLayout()
        layout.addLayout(form)

        t1_cols = [c for c in self.engine.df1.columns if c != "_KEY_"]
        t2_cols = [c for c in self.engine.df2.columns if c != "_KEY_"]

        combos: dict[str, QComboBox] = {}
        for t1 in t1_cols:
            cb = QComboBox()
            cb.addItem("")
            cb.addItems(t2_cols)
            if t1 in self.col_links:
                cb.setCurrentText(self.col_links[t1])
            combos[t1] = cb
            form.addRow(t1, cb)

        row = QHBoxLayout()
        btn_ok = QPushButton("Speichern")
        btn_cancel = QPushButton("Abbrechen")
        row.addWidget(btn_ok)
        row.addWidget(btn_cancel)
        layout.addLayout(row)

        def on_ok():
            self.col_links = {t1: combos[t1].currentText().strip() for t1 in t1_cols if combos[t1].currentText().strip()}
            self.settings.col_links = dict(self.col_links)
            self._save_settings_now()
            dlg.accept()

        btn_ok.clicked.connect(on_ok)
        btn_cancel.clicked.connect(dlg.reject)
        dlg.exec()

    # ---------------- Cuts ----------------
    def open_cuts_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Cuts (Transformationen)")
        layout = QVBoxLayout(dlg)

        cb_split = QCheckBox("Adresse splitten: Straße/Hausnummer trennen")
        cb_phone = QCheckBox("Telefon normalisieren: nur Ziffern (optional +)")
        cb_country = QCheckBox("Country default setzen, wenn leer")
        cb_state = QCheckBox("Bundesland aus PLZ ermitteln")

        cb_split.setChecked(self.cuts.get("split_street_house", True))
        cb_phone.setChecked(self.cuts.get("normalize_phone", True))
        cb_country.setChecked(self.cuts.get("fill_country_default", False))
        cb_state.setChecked(self.cuts.get("infer_state_from_zip", False))

        layout.addWidget(cb_split)
        layout.addWidget(cb_phone)
        layout.addWidget(cb_country)
        layout.addWidget(cb_state)

        btn_country = QPushButton(f"Country-Default (aktuell: {self.country_default_value})")
        layout.addWidget(btn_country)

        btn_ok = QPushButton("Speichern")
        layout.addWidget(btn_ok)

        def on_country():
            v, ok = QInputDialog.getText(self, "Country Default", "Wert für Country (z.B. Deutschland oder DE):")
            if ok and (v or "").strip():
                self.country_default_value = v.strip()
                btn_country.setText(f"Country-Default (aktuell: {self.country_default_value})")

        def on_ok():
            self.cuts["split_street_house"] = cb_split.isChecked()
            self.cuts["normalize_phone"] = cb_phone.isChecked()
            self.cuts["fill_country_default"] = cb_country.isChecked()
            self.cuts["infer_state_from_zip"] = cb_state.isChecked()
            self.settings.cuts = dict(self.cuts)
            self.settings.country_default_value = self.country_default_value
            self._save_settings_now()
            dlg.accept()

        btn_country.clicked.connect(on_country)
        btn_ok.clicked.connect(on_ok)
        dlg.exec()

    # ---------------- Plausible fill row ----------------
    def autofill_current_key_linked(self):
        if not self.engine or self.current_key is None:
            return

        t1_idx = self.engine.t1_row_index_for_key(self.current_key)
        if t1_idx is None:
            return

        t2_df = self.engine.t2_rows_for_key(self.current_key)
        if t2_df is None or len(t2_df) == 0:
            QMessageBox.information(self, "Auto-Fill", "Keine passende Kundennummer in Tabelle 2 gefunden.")
            return

        display_cols = [c for c in self.engine.df1.columns if c != "_KEY_"]
        key_col = self.cb_t1_key.currentText()

        filled = 0

        for t1_col in display_cols:
            if t1_col == key_col:
                continue
            if not is_missing(self.engine.df1.at[t1_idx, t1_col]):
                continue

            t2_col = self.col_links.get(t1_col)
            if not t2_col or t2_col not in t2_df.columns:
                continue

            chosen = None
            for _, r in t2_df.iterrows():
                v = r.get(t2_col)
                if not is_missing(v):
                    chosen = str(v).strip()
                    if chosen:
                        break
            if not chosen:
                continue

            if self.cuts.get("split_street_house", True) and t1_col.lower() in ("street", "straße"):
                street, house = split_street_house(chosen)
                chosen = street
                for cand in ("houseNumber", "hausnummer"):
                    if cand in self.engine.df1.columns and is_missing(self.engine.df1.at[t1_idx, cand]) and house:
                        self.engine.df1.at[t1_idx, cand] = house

            if self.cuts.get("normalize_phone", True) and t1_col.lower() in ("phone", "phonegeneral", "telefon", "festnetz", "mobil", "mobilgeneral"):
                chosen = normalize_phone(chosen)

            self.engine.df1.at[t1_idx, t1_col] = chosen
            filled += 1

        if self.cuts.get("fill_country_default", False) and "country" in self.engine.df1.columns and is_missing(self.engine.df1.at[t1_idx, "country"]):
            self.engine.df1.at[t1_idx, "country"] = self.country_default_value

        if self.cuts.get("infer_state_from_zip", False) and "state" in self.engine.df1.columns and is_missing(self.engine.df1.at[t1_idx, "state"]):
            z = None
            for cand in ("zipCode", "plz", "postalCode"):
                if cand in self.engine.df1.columns:
                    z = self.engine.df1.at[t1_idx, cand]
                    break
            st = state_from_zip_de(z) if z else None
            if st:
                self.engine.df1.at[t1_idx, "state"] = st

        self.show_key(self.current_key)
        QMessageBox.information(self, "Auto-Fill", f"{filled} Felder (Zeile) plausibel gefüllt.")

    # ---------------- Plausible fill all ----------------
    def autofill_all_linked(self):
        if not self.engine:
            QMessageBox.warning(self, "Fehlt", "Bitte erst Start ausführen.")
            return
        if not self.col_links:
            QMessageBox.warning(self, "Fehlt", "Bitte erst Kopplungen definieren.")
            return

        key_col = self.cb_t1_key.currentText()
        total_filled = 0

        for idx, row in self.engine.df1.iterrows():
            key = row.get("_KEY_")
            if not key:
                continue

            t2_df = self.engine.t2_rows_for_key(key)
            if t2_df is None or len(t2_df) == 0:
                continue

            for t1_col, t2_col in self.col_links.items():
                if t1_col in ("_KEY_", key_col):
                    continue
                if t1_col not in self.engine.df1.columns or t2_col not in t2_df.columns:
                    continue
                if not is_missing(self.engine.df1.at[idx, t1_col]):
                    continue

                chosen = None
                for _, r2 in t2_df.iterrows():
                    v = r2.get(t2_col)
                    if not is_missing(v):
                        chosen = str(v).strip()
                        if chosen:
                            break
                if not chosen:
                    continue

                if self.cuts.get("normalize_phone", True) and t1_col.lower() in ("phone", "phonegeneral", "telefon", "festnetz", "mobil", "mobilgeneral"):
                    chosen = normalize_phone(chosen)

                self.engine.df1.at[idx, t1_col] = chosen
                total_filled += 1

                if self.cuts.get("split_street_house", True) and t1_col.lower() in ("street", "straße"):
                    street, house = split_street_house(chosen)
                    self.engine.df1.at[idx, t1_col] = street
                    for cand in ("houseNumber", "hausnummer"):
                        if cand in self.engine.df1.columns and is_missing(self.engine.df1.at[idx, cand]) and house:
                            self.engine.df1.at[idx, cand] = house

            if self.cuts.get("fill_country_default", False) and "country" in self.engine.df1.columns and is_missing(self.engine.df1.at[idx, "country"]):
                self.engine.df1.at[idx, "country"] = self.country_default_value

            if self.cuts.get("infer_state_from_zip", False) and "state" in self.engine.df1.columns and is_missing(self.engine.df1.at[idx, "state"]):
                z = None
                for cand in ("zipCode", "plz", "postalCode"):
                    if cand in self.engine.df1.columns:
                        z = self.engine.df1.at[idx, cand]
                        break
                st = state_from_zip_de(z) if z else None
                if st:
                    self.engine.df1.at[idx, "state"] = st

        if self.current_key is not None:
            self.show_key(self.current_key)

        QMessageBox.information(self, "Auto-Fill Gesamt", f"{total_filled} Zellen in der gesamten Tabelle gefüllt.")

    # ---------------- Add column ----------------
    def add_column_t1_global(self):
        if not self.engine or not self.t1:
            QMessageBox.warning(self, "Fehlt", "Bitte zuerst Tabelle 1 laden und Start ausführen.")
            return

        name, ok = QInputDialog.getText(self, "Neue Spalte", "Name der neuen Spalte (Tabelle 1):")
        if not ok:
            return
        name = (name or "").strip()
        if not name:
            QMessageBox.warning(self, "Ungültig", "Spaltenname ist leer.")
            return
        if name == "_KEY_":
            QMessageBox.warning(self, "Ungültig", "Spaltenname '_KEY_' ist reserviert.")
            return
        if name in self.engine.df1.columns:
            QMessageBox.information(self, "Info", f"Spalte '{name}' existiert bereits.")
            return

        default_value, ok2 = QInputDialog.getText(self, "Default-Wert", "Optional: Default-Wert für alle Zeilen (leer = leer):")
        if not ok2:
            return

        self.engine.df1[name] = default_value or ""

        if self.current_key is not None:
            self.show_key(self.current_key)

        QMessageBox.information(self, "OK", f"Spalte '{name}' wurde hinzugefügt.")
