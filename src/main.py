# -*- coding: utf-8 -*-
"""
Oblik Inventory
===============
Перша робоча desktop-версія програми обліку майна.

Ключові принципи:
- Excel залишається зрозумілим користувачу носієм/експортом.
- pandas використовується для аналізу даних.
- openpyxl використовується для читання/запису Excel зі збереженням структури.
- PySide6 забезпечує Windows-інтерфейс.
- зміни НЕ блокуються: програма попереджає і журналює їх;
- "Штатна потреба" не перераховується і переноситься як є;
- "Поточний стан" автоматично перебудовується тільки для поштучного майна,
  яке можна однозначно ідентифікувати за інвентарним або заводським номером.
  Запаси без такого ідентифікатора поки не згортаються автоматично.
"""

from __future__ import annotations

import getpass
import sys
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QAction, QColor, QFont
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QDialogButtonBox, QFileDialog,
    QFormLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget,
    QMainWindow, QMessageBox, QPushButton, QSplitter, QStatusBar,
    QTableWidget, QTableWidgetItem, QTextEdit, QToolBar, QVBoxLayout, QWidget,
)

APP_TITLE = "Oblik Inventory"
APP_VERSION = "0.1.0"

SHEET_STAFF = "Штат"
SHEET_MOVEMENT = "Рух майна"
SHEET_CURRENT = "Поточний стан"
SHEET_SUMMARY = "Зведений"
SHEET_CHANGES = "Контроль змін"
REQUIRED_SHEETS = [SHEET_STAFF, SHEET_MOVEMENT, SHEET_CURRENT, SHEET_SUMMARY, SHEET_CHANGES]

MAIN_HEADERS = [
    "№ з/п",
    "Підства \n(номер наряду\n/договір)",
    "Дата підстави \n(наряду\n/договору)",
    "Номер докумету про отримання \n(акт п/п, видаткова накладна)",
    "Дата документу \nпро отримання",
    "Походження \nмайна",
    "Діапазони частот на яких працює РЕБ",
    "Номенклатурний \nкод",
    "Тип майна",
    "Інвентарний номер",
    "Узагальнанена назва згідно номенклатору",
    "Узагальнена назва",
    "Найменування майна згідно документів",
    "Номер матеріалу в SAP",
    "Штатна потреба",
    "Ціна",
    "Кількість",
    "Сума",
    "Заводський \nномер",
    "Окремий підрозді \nбригади де знаходиться на даний момент",
    "Окремий підрозділ \nбатальйону, дивізіону\n де знаходиться на даний момент",
    "Номер документу згідно\n якого підрозділ це \nотримав на даний момент",
    "Дата документу згідно\n якого знаходиться в підрозділі на даний момент",
    "Справний/Несправний \n(переданий в ремонт)/Знищений/Втрачений",
    "Дата якщо пошкоджений/\nвтрачений/знищений",
    "Якщо БПВ або знищений\nномер наказу \nна призначення СР",
    "Дата наказу на \nпризначення СР",
    "Номер наказу \nна результати СР",
    "Дата наказу \nна результати СР",
    "Номер наказу на списання",
    "Дата наказу на списання",
    "Номер Єдиного акту списання",
    "Дата єдиного акту списання",
    "Примітка",
]

CHANGE_HEADERS = [
    "№ з/п", "Дата і час", "Тип активу", "Таблиця", "Інв. №",
    "Найменування майна", "Поле", "Було написано", "Стало написано",
    "Тип зміни", "Причина", "Користувач",
]

STAFF_HEADERS = [
    "№ з/п", "Номенклатурний код", "Найменування згідно номенклатору",
    "Узагальнена кількість за військову частину", "", "",
    "№ з/п", "Номенклатурний код", "Найменування згідно номенклатору",
    "Кількість", "Окремий підрозділ бригади", "Окремий підрозділ батальйону/дивізіону",
]

DUPLICATE_FIELDS = {
    "order_no": MAIN_HEADERS[1],
    "order_date": MAIN_HEADERS[2],
    "receive_doc": MAIN_HEADERS[3],
    "receive_date": MAIN_HEADERS[4],
    "inventory_no": MAIN_HEADERS[9],
    "serial_no": MAIN_HEADERS[18],
}


def norm(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return " ".join(str(value).strip().casefold().split())


def display_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    return str(value)


def parse_user_value(header: str, text: str) -> Any:
    text = text.strip()
    if not text:
        return None
    h = header.casefold()
    if "дата" in h:
        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass
        return text
    if header in ("Ціна", "Кількість", "Сума", "Штатна потреба"):
        candidate = text.replace(" ", "").replace(",", ".")
        try:
            return float(candidate)
        except ValueError:
            return text
    return text


@dataclass
class DuplicateInfo:
    score: int = 0
    other_excel_rows: tuple[int, ...] = ()


class OblikWorkbook:
    """Excel-книга + основна логіка обліку."""

    def __init__(self) -> None:
        self.path: Optional[Path] = None
        self.wb = None
        self.dirty = False

    def create_new(self, path: Path) -> None:
        wb = Workbook()
        wb.remove(wb.active)
        self._create_main_sheet(wb, SHEET_MOVEMENT)
        self._create_main_sheet(wb, SHEET_CURRENT)

        ws_staff = wb.create_sheet(SHEET_STAFF)
        for col, header in enumerate(STAFF_HEADERS, 1):
            ws_staff.cell(1, col, header)
        self._style_header(ws_staff, 1, len(STAFF_HEADERS))
        ws_staff.freeze_panes = "A2"

        ws_summary = wb.create_sheet(SHEET_SUMMARY)
        summary_headers = ["№ з/п", "Найменування ОВТ", "За штатом", "За списком", "Наявні", "Несправні", "Потреба", "БПВ", "Забезпеченість", "Примітка", "Дані"]
        for col, header in enumerate(summary_headers, 1):
            ws_summary.cell(1, col, header)
        self._style_header(ws_summary, 1, len(summary_headers))

        ws_changes = wb.create_sheet(SHEET_CHANGES)
        for col, header in enumerate(CHANGE_HEADERS, 1):
            ws_changes.cell(1, col, header)
        self._style_header(ws_changes, 1, len(CHANGE_HEADERS))
        ws_changes.freeze_panes = "A2"

        wb.save(path)
        self.load(path)

    def _create_main_sheet(self, wb: Workbook, name: str) -> None:
        ws = wb.create_sheet(name)
        for col, header in enumerate(MAIN_HEADERS, 1):
            ws.cell(1, col, header)
            ws.cell(2, col, col)
        self._style_header(ws, 1, len(MAIN_HEADERS))
        for cell in ws[2]:
            cell.font = Font(bold=True, color="666666")
            cell.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(MAIN_HEADERS))}1"
        for idx in range(1, len(MAIN_HEADERS) + 1):
            ws.column_dimensions[get_column_letter(idx)].width = 18

    def _style_header(self, ws, row: int, count: int) -> None:
        fill = PatternFill("solid", fgColor="1F4E78")
        side = Side(style="thin", color="A0A0A0")
        for col in range(1, count + 1):
            cell = ws.cell(row, col)
            cell.fill = fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=side, right=side, top=side, bottom=side)
        ws.row_dimensions[row].height = 60

    def load(self, path: Path) -> None:
        self.wb = load_workbook(path)
        missing = [s for s in REQUIRED_SHEETS if s not in self.wb.sheetnames]
        if missing:
            raise ValueError("У книзі відсутні аркуші: " + ", ".join(missing))
        self.path = path
        self.dirty = False

    def save(self, path: Optional[Path] = None) -> None:
        if self.wb is None:
            raise RuntimeError("Книгу не відкрито")
        target = path or self.path
        if not target:
            raise RuntimeError("Не вказано шлях для збереження")
        self.wb.save(target)
        self.path = Path(target)
        self.dirty = False

    def headers(self, sheet_name: str) -> list[str]:
        ws = self.wb[sheet_name]
        return [str(ws.cell(1, c).value or "") for c in range(1, ws.max_column + 1)]

    def _row_has_data(self, ws, row: int, start_col: int = 1, end_col: Optional[int] = None) -> bool:
        end_col = end_col or ws.max_column
        return any(ws.cell(row, c).value not in (None, "") for c in range(start_col, end_col + 1))

    def _first_empty_data_row(self, ws, start_row: int = 3) -> int:
        upper = max(ws.max_row + 2, start_row + 2)
        for row in range(start_row, upper + 1):
            if not self._row_has_data(ws, row, 1, ws.max_column):
                return row
        return upper + 1

    def dataframe(self, sheet_name: str) -> pd.DataFrame:
        ws = self.wb[sheet_name]
        headers = self.headers(sheet_name)
        rows = []
        start_row = 3 if sheet_name in (SHEET_MOVEMENT, SHEET_CURRENT) else 2
        for row_no in range(start_row, ws.max_row + 1):
            values = [ws.cell(row_no, col).value for col in range(1, len(headers) + 1)]
            if not any(value not in (None, "") for value in values):
                continue
            row = {headers[i]: values[i] for i in range(len(headers))}
            row["_excel_row"] = row_no
            rows.append(row)
        return pd.DataFrame(rows)

    def append_record(self, sheet_name: str, values: dict[str, Any], reason: str = "") -> int:
        ws = self.wb[sheet_name]
        headers = self.headers(sheet_name)
        start_row = 3 if sheet_name in (SHEET_MOVEMENT, SHEET_CURRENT) else 2
        row = self._first_empty_data_row(ws, start_row)
        template_row = start_row

        for col in range(1, len(headers) + 1):
            if template_row <= ws.max_row:
                src = ws.cell(template_row, col)
                dst = ws.cell(row, col)
                if src.has_style:
                    dst._style = copy(src._style)

        if headers and headers[0] == "№ з/п":
            values[headers[0]] = self.next_sequence(sheet_name)

        for col, header in enumerate(headers, 1):
            ws.cell(row, col, values.get(header))

        self.dirty = True
        if sheet_name == SHEET_MOVEMENT:
            self._log_change(values, sheet_name, "Весь запис", "", self._record_summary(values), "Додавання", reason)
        return row

    def update_record(self, sheet_name: str, excel_row: int, new_values: dict[str, Any], reason: str = ""):
        ws = self.wb[sheet_name]
        headers = self.headers(sheet_name)
        old_values = {header: ws.cell(excel_row, i + 1).value for i, header in enumerate(headers)}
        changes = []

        for i, header in enumerate(headers, 1):
            if header == "№ з/п":
                continue
            old = old_values.get(header)
            new = new_values.get(header)
            if norm(old) != norm(new):
                ws.cell(excel_row, i, new)
                changes.append((header, old, new))

        if changes:
            self.dirty = True
            merged = dict(old_values)
            merged.update(new_values)
            for field, old, new in changes:
                self._log_change(merged, sheet_name, field, old, new, "Виправлення", reason)
        return changes

    def delete_record(self, sheet_name: str, excel_row: int, reason: str = "") -> None:
        ws = self.wb[sheet_name]
        headers = self.headers(sheet_name)
        old_values = {header: ws.cell(excel_row, i + 1).value for i, header in enumerate(headers)}
        summary = self._record_summary(old_values)
        for col in range(1, len(headers) + 1):
            ws.cell(excel_row, col).value = None
        self.dirty = True
        self._log_change(old_values, sheet_name, "Весь запис", summary, "", "Видалення", reason)

    def next_sequence(self, sheet_name: str) -> int:
        ws = self.wb[sheet_name]
        start_row = 3 if sheet_name in (SHEET_MOVEMENT, SHEET_CURRENT) else 2
        values = []
        for row in range(start_row, ws.max_row + 1):
            try:
                values.append(int(ws.cell(row, 1).value))
            except (TypeError, ValueError):
                pass
        return max(values, default=0) + 1

    def _record_summary(self, values: dict[str, Any]) -> str:
        keys = [MAIN_HEADERS[9], MAIN_HEADERS[18], MAIN_HEADERS[12], MAIN_HEADERS[16], MAIN_HEADERS[19], MAIN_HEADERS[20]]
        parts = []
        for key in keys:
            value = values.get(key)
            if value not in (None, ""):
                parts.append(f"{key.replace(chr(10), ' ')}={display_value(value)}")
        return "; ".join(parts)[:1000]

    def _log_change(self, record, table, field, old, new, change_type, reason):
        ws = self.wb[SHEET_CHANGES]
        row = 2
        while row <= max(ws.max_row + 1, 2):
            if not any(ws.cell(row, col).value not in (None, "") for col in range(2, 13)):
                break
            row += 1

        seq = 1
        for r in range(2, row):
            try:
                seq = max(seq, int(ws.cell(r, 1).value or 0) + 1)
            except (TypeError, ValueError):
                pass

        vals = [
            seq, datetime.now(), record.get(MAIN_HEADERS[8], ""), table,
            record.get(MAIN_HEADERS[9], ""),
            record.get(MAIN_HEADERS[12], "") or record.get(MAIN_HEADERS[11], ""),
            field, display_value(old), display_value(new), change_type,
            reason, getpass.getuser(),
        ]
        for col, value in enumerate(vals, 1):
            ws.cell(row, col, value)
        ws.cell(row, 2).number_format = "dd.mm.yyyy hh:mm:ss"

    def duplicate_map(self) -> dict[int, DuplicateInfo]:
        df = self.dataframe(SHEET_MOVEMENT)
        if df.empty:
            return {}

        records = []
        for _, row in df.iterrows():
            identifier = norm(row.get(DUPLICATE_FIELDS["inventory_no"])) or norm(row.get(DUPLICATE_FIELDS["serial_no"]))
            sig = [
                norm(row.get(DUPLICATE_FIELDS["order_no"])),
                norm(row.get(DUPLICATE_FIELDS["order_date"])),
                norm(row.get(DUPLICATE_FIELDS["receive_doc"])),
                norm(row.get(DUPLICATE_FIELDS["receive_date"])),
                identifier,
            ]
            records.append((int(row["_excel_row"]), sig))

        inverted = [defaultdict(set) for _ in range(5)]
        for excel_row, sig in records:
            for i, value in enumerate(sig):
                if value:
                    inverted[i][value].add(excel_row)

        result = {}
        for excel_row, sig in records:
            counts = defaultdict(int)
            for i, value in enumerate(sig):
                if not value:
                    continue
                for other in inverted[i].get(value, set()):
                    if other != excel_row:
                        counts[other] += 1
            if counts:
                max_score = max(counts.values())
                others = tuple(sorted(k for k, score in counts.items() if score == max_score))
                result[excel_row] = DuplicateInfo(max_score, others)
            else:
                result[excel_row] = DuplicateInfo()
        return result

    def score_candidate_duplicate(self, candidate, ignore_excel_row=None) -> DuplicateInfo:
        df = self.dataframe(SHEET_MOVEMENT)
        if df.empty:
            return DuplicateInfo()

        candidate_sig = [
            norm(candidate.get(DUPLICATE_FIELDS["order_no"])),
            norm(candidate.get(DUPLICATE_FIELDS["order_date"])),
            norm(candidate.get(DUPLICATE_FIELDS["receive_doc"])),
            norm(candidate.get(DUPLICATE_FIELDS["receive_date"])),
            norm(candidate.get(DUPLICATE_FIELDS["inventory_no"])) or norm(candidate.get(DUPLICATE_FIELDS["serial_no"])),
        ]

        best = 0
        rows = []
        for _, row in df.iterrows():
            rnum = int(row["_excel_row"])
            if ignore_excel_row and rnum == ignore_excel_row:
                continue
            sig = [
                norm(row.get(DUPLICATE_FIELDS["order_no"])),
                norm(row.get(DUPLICATE_FIELDS["order_date"])),
                norm(row.get(DUPLICATE_FIELDS["receive_doc"])),
                norm(row.get(DUPLICATE_FIELDS["receive_date"])),
                norm(row.get(DUPLICATE_FIELDS["inventory_no"])) or norm(row.get(DUPLICATE_FIELDS["serial_no"])),
            ]
            score = sum(1 for a, b in zip(candidate_sig, sig) if a and b and a == b)
            if score > best:
                best = score
                rows = [rnum]
            elif score == best and score > 0:
                rows.append(rnum)
        return DuplicateInfo(best, tuple(rows))

    def rebuild_current_state(self) -> tuple[int, int]:
        """Поточний стан по поштучному майну. Штатну потребу не перераховує."""
        df = self.dataframe(SHEET_MOVEMENT)
        ws_cur = self.wb[SHEET_CURRENT]
        if df.empty:
            return 0, 0

        headers = self.headers(SHEET_MOVEMENT)
        name_col = MAIN_HEADERS[12]
        inv_col = MAIN_HEADERS[9]
        serial_col = MAIN_HEADERS[18]
        movement_date_col = MAIN_HEADERS[22]
        receive_date_col = MAIN_HEADERS[4]

        def key_for(row) -> str:
            inv = norm(row.get(inv_col))
            if inv:
                return "INV|" + inv
            serial = norm(row.get(serial_col))
            if serial:
                return "SER|" + serial + "|" + norm(row.get(name_col))
            return ""

        df = df.copy()
        df["_asset_key"] = df.apply(key_for, axis=1)
        identified = df[df["_asset_key"] != ""].copy()
        skipped = int((df["_asset_key"] == "").sum())

        def as_ts(value):
            if value is None or value == "":
                return pd.NaT
            return pd.to_datetime(value, dayfirst=True, errors="coerce")

        identified["_sort_date"] = identified[movement_date_col].map(as_ts)
        identified["_receive_date"] = identified[receive_date_col].map(as_ts)
        identified["_sort_date"] = identified["_sort_date"].fillna(identified["_receive_date"])
        identified["_sort_date"] = identified["_sort_date"].fillna(pd.Timestamp("1900-01-01"))
        identified = identified.sort_values(["_asset_key", "_sort_date", "_excel_row"])
        latest = identified.groupby("_asset_key", as_index=False).tail(1)

        for row in range(3, ws_cur.max_row + 1):
            for col in range(1, min(len(headers), ws_cur.max_column) + 1):
                ws_cur.cell(row, col).value = None

        out_row = 3
        for seq, (_, record) in enumerate(latest.iterrows(), 1):
            for col, header in enumerate(headers, 1):
                value = record.get(header)
                if header == "№ з/п":
                    value = seq
                ws_cur.cell(out_row, col, value)
            out_row += 1

        self.dirty = True
        return len(latest), skipped


class RecordDialog(QDialog):
    """Діалог усіх полів одного запису."""

    def __init__(self, headers, values=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Запис майна")
        self.resize(760, 760)
        self.inputs = {}

        outer = QVBoxLayout(self)
        scroll_host = QWidget()
        form = QFormLayout(scroll_host)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignTop)

        for header in headers:
            if header == "№ з/п":
                continue
            current = values.get(header) if values else None
            label = QLabel(header.replace("\n", " "))
            label.setWordWrap(True)

            if header == MAIN_HEADERS[8]:
                widget = QComboBox()
                widget.setEditable(True)
                widget.addItems(["", "Необоротний актив", "Запас"])
                widget.setCurrentText(display_value(current))
            elif header == MAIN_HEADERS[23]:
                widget = QComboBox()
                widget.setEditable(True)
                widget.addItems(["", "Справний", "Несправний", "Переданий в ремонт", "Знищений", "Втрачений"])
                widget.setCurrentText(display_value(current))
            elif header == "Примітка":
                widget = QTextEdit()
                widget.setMaximumHeight(90)
                widget.setPlainText(display_value(current))
            else:
                widget = QLineEdit(display_value(current))
                if "дата" in header.casefold():
                    widget.setPlaceholderText("дд.мм.рррр")

            self.inputs[header] = widget
            form.addRow(label, widget)

        from PySide6.QtWidgets import QScrollArea
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(scroll_host)
        outer.addWidget(scroll)

        self.reason = QLineEdit()
        self.reason.setPlaceholderText("Причина зміни / уточнення (за потреби)")
        outer.addWidget(QLabel("Причина зміни:"))
        outer.addWidget(self.reason)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        outer.addWidget(buttons)

    def data(self):
        out = {}
        for header, widget in self.inputs.items():
            if isinstance(widget, QComboBox):
                text = widget.currentText()
            elif isinstance(widget, QTextEdit):
                text = widget.toPlainText()
            else:
                text = widget.text()
            out[header] = parse_user_value(header, text)
        return out


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.model = OblikWorkbook()
        self.current_sheet = SHEET_MOVEMENT
        self.setWindowTitle(f"{APP_TITLE} {APP_VERSION}")
        self.resize(1500, 900)
        self._build_ui()
        self._apply_style()

    def _build_ui(self):
        toolbar = QToolBar("Основні дії")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        actions = [
            ("Відкрити Excel", self.open_file),
            ("Нова книга", self.new_file),
            ("Зберегти", self.save_file),
            ("Зберегти як", self.save_as),
        ]
        for title, handler in actions:
            action = QAction(title, self)
            action.triggered.connect(handler)
            toolbar.addAction(action)

        toolbar.addSeparator()
        rebuild = QAction("Оновити поточний стан", self)
        rebuild.triggered.connect(self.rebuild_current)
        toolbar.addAction(rebuild)

        root = QWidget()
        layout = QVBoxLayout(root)
        top = QHBoxLayout()

        self.search = QLineEdit()
        self.search.setPlaceholderText("Пошук у відкритій таблиці…")
        self.search.textChanged.connect(self.apply_filter)
        top.addWidget(QLabel("Пошук:"))
        top.addWidget(self.search, 1)

        self.btn_add = QPushButton("+ Додати запис")
        self.btn_add.clicked.connect(self.add_record)
        self.btn_edit = QPushButton("Редагувати")
        self.btn_edit.clicked.connect(self.edit_record)
        self.btn_delete = QPushButton("Видалити")
        self.btn_delete.clicked.connect(self.delete_record)
        top.addWidget(self.btn_add)
        top.addWidget(self.btn_edit)
        top.addWidget(self.btn_delete)
        layout.addLayout(top)

        splitter = QSplitter()
        self.sidebar = QListWidget()
        self.sidebar.addItems(REQUIRED_SHEETS)
        self.sidebar.setMaximumWidth(220)
        self.sidebar.currentTextChanged.connect(self.select_sheet)
        self.sidebar.setCurrentRow(1)
        splitter.addWidget(self.sidebar)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.doubleClicked.connect(self.edit_record)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setDefaultSectionSize(160)
        splitter.addWidget(self.table)
        splitter.setStretchFactor(1, 1)
        layout.addWidget(splitter, 1)

        self.setCentralWidget(root)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Відкрийте існуючу книгу або створіть нову")
        self._update_edit_permissions()

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow { background: #f4f6f8; }
            QToolBar { background: #17365d; color: white; spacing: 8px; padding: 5px; }
            QToolButton { color: white; padding: 6px 10px; }
            QListWidget { background: #203864; color: white; border: none; font-size: 14px; }
            QListWidget::item { padding: 13px 12px; }
            QListWidget::item:selected { background: #2f75b5; font-weight: bold; }
            QPushButton { padding: 7px 12px; border-radius: 4px; background: #2f75b5; color: white; }
            QPushButton:disabled { background: #aab4be; }
            QLineEdit, QComboBox, QTextEdit { padding: 6px; background: white; border: 1px solid #c7cdd4; border-radius: 3px; }
            QTableWidget { background: white; gridline-color: #dde2e7; }
            QHeaderView::section { background: #d9eaf7; padding: 6px; border: 1px solid #b8c9d8; font-weight: bold; }
        """)

    def _maybe_save(self):
        if not self.model.dirty:
            return True
        answer = QMessageBox.question(self, "Незбережені зміни", "Зберегти зміни перед продовженням?", QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        if answer == QMessageBox.Cancel:
            return False
        if answer == QMessageBox.Yes:
            self.save_file()
            return not self.model.dirty
        return True

    def open_file(self):
        if not self._maybe_save():
            return
        path, _ = QFileDialog.getOpenFileName(self, "Відкрити книгу обліку", "", "Excel (*.xlsx *.xlsm)")
        if not path:
            return
        try:
            self.model.load(Path(path))
            self.refresh_table()
            self.statusBar().showMessage(f"Відкрито: {path}")
        except Exception as exc:
            QMessageBox.critical(self, "Помилка", str(exc))

    def new_file(self):
        if not self._maybe_save():
            return
        path, _ = QFileDialog.getSaveFileName(self, "Створити книгу", "Oblik.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            self.model.create_new(Path(path))
            self.refresh_table()
        except Exception as exc:
            QMessageBox.critical(self, "Помилка", str(exc))

    def save_file(self):
        if self.model.wb is None:
            return
        try:
            self.model.save()
            self.statusBar().showMessage(f"Збережено: {self.model.path}", 5000)
        except PermissionError:
            QMessageBox.warning(self, "Файл зайнятий", "Закрийте книгу в Excel і повторіть.")
        except Exception as exc:
            QMessageBox.critical(self, "Помилка збереження", str(exc))

    def save_as(self):
        if self.model.wb is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти копію", "Oblik_copy.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            self.model.save(Path(path))
            self.refresh_table()
        except Exception as exc:
            QMessageBox.critical(self, "Помилка", str(exc))

    def select_sheet(self, name):
        if not name:
            return
        self.current_sheet = name
        self.search.clear()
        self.refresh_table()
        self._update_edit_permissions()

    def _update_edit_permissions(self):
        editable = self.current_sheet in (SHEET_MOVEMENT, SHEET_STAFF)
        enabled = editable and self.model.wb is not None
        self.btn_add.setEnabled(enabled)
        self.btn_edit.setEnabled(enabled)
        self.btn_delete.setEnabled(enabled)

    def refresh_table(self):
        self.table.clear()
        if self.model.wb is None or self.current_sheet not in self.model.wb.sheetnames:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self._update_edit_permissions()
            return

        df = self.model.dataframe(self.current_sheet)
        headers = self.model.headers(self.current_sheet)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels([header.replace("\n", " ") for header in headers])

        if df.empty:
            self.table.setRowCount(0)
            self._update_edit_permissions()
            return

        data_headers = [c for c in df.columns if c != "_excel_row"]
        self.table.setRowCount(len(df))
        duplicate_map = self.model.duplicate_map() if self.current_sheet == SHEET_MOVEMENT else {}

        for r_idx, (_, row) in enumerate(df.iterrows()):
            excel_row = int(row["_excel_row"])
            dup = duplicate_map.get(excel_row, DuplicateInfo())
            for c_idx, header in enumerate(data_headers):
                item = QTableWidgetItem(display_value(row.get(header)))
                item.setData(Qt.UserRole, excel_row)
                if self.current_sheet == SHEET_MOVEMENT:
                    if dup.score >= 5:
                        item.setForeground(QColor("#c00000"))
                        item.setFont(QFont(item.font().family(), item.font().pointSize(), QFont.Bold))
                    elif dup.score == 4:
                        item.setForeground(QColor("#e26b0a"))
                    elif dup.score == 3:
                        item.setForeground(QColor("#9c6500"))
                    if dup.score >= 3:
                        item.setToolTip(f"Збіг {dup.score}/5 з Excel-рядками: {', '.join(map(str, dup.other_excel_rows))}. Збереження не блокується.")
                self.table.setItem(r_idx, c_idx, item)

        self.table.resizeRowsToContents()
        self.statusBar().showMessage(f"{self.current_sheet}: {len(df)} записів | {self.model.path or ''}")
        self.apply_filter()
        self._update_edit_permissions()

    def apply_filter(self):
        query = norm(self.search.text())
        for row in range(self.table.rowCount()):
            visible = True
            if query:
                visible = any(query in norm(self.table.item(row, col).text() if self.table.item(row, col) else "") for col in range(self.table.columnCount()))
            self.table.setRowHidden(row, not visible)

    def _selected_excel_row(self):
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return int(item.data(Qt.UserRole)) if item and item.data(Qt.UserRole) else None

    def add_record(self):
        if self.model.wb is None or self.current_sheet not in (SHEET_MOVEMENT, SHEET_STAFF):
            return
        headers = self.model.headers(self.current_sheet)
        dialog = RecordDialog(headers, parent=self)
        if dialog.exec() != QDialog.Accepted:
            return
        data = dialog.data()

        if self.current_sheet == SHEET_MOVEMENT:
            dup = self.model.score_candidate_duplicate(data)
            if dup.score >= 3 and not self._duplicate_confirmation(dup):
                return
            asset_type = norm(data.get(MAIN_HEADERS[8]))
            if "необорот" in asset_type and not norm(data.get(MAIN_HEADERS[9])):
                answer = QMessageBox.warning(self, "Інвентарний номер не вказаний", "Для необоротного активу інвентарний номер порожній. Зберегти запис все одно?", QMessageBox.Yes | QMessageBox.No)
                if answer != QMessageBox.Yes:
                    return

        self.model.append_record(self.current_sheet, data, dialog.reason.text())
        self.refresh_table()

    def edit_record(self, *_):
        if self.model.wb is None or self.current_sheet not in (SHEET_MOVEMENT, SHEET_STAFF):
            return
        excel_row = self._selected_excel_row()
        if not excel_row:
            QMessageBox.information(self, "Редагування", "Виберіть рядок.")
            return
        ws = self.model.wb[self.current_sheet]
        headers = self.model.headers(self.current_sheet)
        values = {header: ws.cell(excel_row, i + 1).value for i, header in enumerate(headers)}
        dialog = RecordDialog(headers, values, self)
        if dialog.exec() != QDialog.Accepted:
            return
        data = dialog.data()
        if self.current_sheet == SHEET_MOVEMENT:
            dup = self.model.score_candidate_duplicate(data, excel_row)
            if dup.score >= 3 and not self._duplicate_confirmation(dup):
                return
        changes = self.model.update_record(self.current_sheet, excel_row, data, dialog.reason.text())
        if changes:
            self.refresh_table()
            self.statusBar().showMessage(f"Змінено полів: {len(changes)}. Записано в Контроль змін.", 5000)

    def delete_record(self):
        if self.model.wb is None or self.current_sheet not in (SHEET_MOVEMENT, SHEET_STAFF):
            return
        excel_row = self._selected_excel_row()
        if not excel_row:
            return
        reason, ok = self._ask_reason("Причина видалення")
        if not ok:
            return
        answer = QMessageBox.warning(self, "Видалення", "Запис буде прибрано з таблиці, але інформація про видалення залишиться у 'Контроль змін'. Продовжити?", QMessageBox.Yes | QMessageBox.No)
        if answer != QMessageBox.Yes:
            return
        self.model.delete_record(self.current_sheet, excel_row, reason)
        self.refresh_table()

    def _ask_reason(self, title):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QVBoxLayout(dialog)
        edit = QLineEdit()
        edit.setPlaceholderText("Наприклад: дубль, уточнення документів, помилка минулих років")
        layout.addWidget(edit)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        ok = dialog.exec() == QDialog.Accepted
        return edit.text(), ok

    def _duplicate_confirmation(self, dup):
        if dup.score >= 5:
            level = "ПОВНИЙ ДУБЛЬ 5/5"
        elif dup.score == 4:
            level = "ДУЖЕ СХОЖИЙ ЗАПИС 4/5"
        else:
            level = "МОЖЛИВИЙ ДУБЛЬ 3/5"
        text = f"{level}\n\nСхожі Excel-рядки: {', '.join(map(str, dup.other_excel_rows))}\n\nСистема не блокує зміну. Зберегти все одно?"
        return QMessageBox.warning(self, "Контроль дублювання", text, QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes

    def rebuild_current(self):
        if self.model.wb is None:
            return
        try:
            count, skipped = self.model.rebuild_current_state()
            self.refresh_table()
            QMessageBox.information(
                self, "Поточний стан оновлено",
                f"Сформовано поштучних записів: {count}.\n"
                f"Записів без інвентарного/заводського номера, які не згорнуті автоматично: {skipped}.\n\n"
                "Колонка 'Штатна потреба' не перераховувалась — значення перенесені як є.",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Помилка розрахунку", str(exc))

    def closeEvent(self, event):
        if self._maybe_save():
            event.accept()
        else:
            event.ignore()


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    window = MainWindow()
    window.show()

    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent
    for candidate in (base / "Oblik.xlsx", base / "data" / "Oblik.xlsx"):
        if candidate.exists():
            QTimer.singleShot(0, lambda p=candidate: (window.model.load(p), window.refresh_table()))
            break

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
