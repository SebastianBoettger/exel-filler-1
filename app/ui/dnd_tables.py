from __future__ import annotations

from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QMenu, QInputDialog
from PySide6.QtCore import Qt, QMimeData
from PySide6.QtGui import QDrag, QAction

MIME = "application/x-excel-filler-cell"


def _is_missing_local(v: str | None) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    if s == "":
        return True
    return s.lower() in {"nan", "none", "null", "-", "n/a"}


class SourceTable(QTableWidget):
    """Rechts: Drag-Quelle (auch Multi-Selection)"""

    def startDrag(self, supportedActions):
        # Sammle selektierte Zellen (in Visual-Reihenfolge: row, col)
        items = self.selectedItems()
        if not items:
            item = self.currentItem()
            if not item:
                return
            items = [item]

        # Sortieren nach row/col für deterministisches Zusammenfügen
        cells = sorted([(it.row(), it.column(), it.text()) for it in items], key=lambda x: (x[0], x[1]))
        parts = [t for _, _, t in cells if not _is_missing_local(t)]
        text = "\n".join(parts)  # default: Zeilenweise

        mime = QMimeData()
        mime.setData(MIME, text.encode("utf-8"))

        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.CopyAction)


class TargetTable(QTableWidget):
    """Links: Drop-Ziel (Replace/Append Menü)"""

    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat(MIME):
            e.acceptProposedAction()
        else:
            super().dragEnterEvent(e)

    def dragMoveEvent(self, e):
        if e.mimeData().hasFormat(MIME):
            e.acceptProposedAction()
        else:
            super().dragMoveEvent(e)

    def dropEvent(self, e):
        if not e.mimeData().hasFormat(MIME):
            return super().dropEvent(e)

        text = bytes(e.mimeData().data(MIME)).decode("utf-8")

        idx = self.indexAt(e.position().toPoint())
        if not idx.isValid():
            e.ignore()
            return

        row, col = idx.row(), idx.column()

        existing_item = self.item(row, col)
        existing_text = existing_item.text() if existing_item else ""
        existing_is_missing = _is_missing_local(existing_text)

        # Wenn Ziel belegt -> Menü
        if not existing_is_missing:
            menu = QMenu(self)

            act_replace = QAction("Ersetzen", self)
            menu.addAction(act_replace)

            sub = menu.addMenu("Anhängen")
            separators = ["/", ", ", "; ", ",", ";", " | ", "\n"]
            sep_actions = []
            for sep in separators:
                a = QAction(f"mit Separator: {repr(sep)}", self)
                sub.addAction(a)
                sep_actions.append((a, sep))

            sub.addSeparator()
            act_custom = QAction("eigener Separator…", self)
            sub.addAction(act_custom)

            chosen_action = menu.exec(self.viewport().mapToGlobal(e.position().toPoint()))
            if chosen_action is None:
                e.ignore()
                return

            if chosen_action == act_replace:
                new_text = text
            elif chosen_action == act_custom:
                sep, ok = QInputDialog.getText(self, "Separator", "Separator eingeben (z.B. '/' oder '; '):")
                if not ok:
                    e.ignore()
                    return
                new_text = f"{existing_text}{sep}{text}"
            else:
                sep = None
                for a, s in sep_actions:
                    if chosen_action == a:
                        sep = s
                        break
                if sep is None:
                    e.ignore()
                    return
                new_text = f"{existing_text}{sep}{text}"
        else:
            new_text = text

        # Schreiben
        self.blockSignals(True)
        item = self.item(row, col)
        if item is None:
            item = QTableWidgetItem()
            self.setItem(row, col, item)
        item.setText(new_text)
        self.blockSignals(False)

        self.setCurrentCell(row, col)
        self.editItem(self.item(row, col))

        mw = self.window()
        if hasattr(mw, "on_t1_cell_dropped"):
            mw.on_t1_cell_dropped(row, col, new_text)

        e.acceptProposedAction()
