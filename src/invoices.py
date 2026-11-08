"""Шаблон Excel-накладної та перевірка її даних перед імпортом."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


TEMPLATE_SHEET = "Накладна"
INVOICE_SHEET = "Накладні"
ITEM_HEADERS = [
    "№ з/п", "Номенклатурний код", "Найменування майна", "Од. виміру",
    "Тип майна", "Кількість", "Ціна", "Інвентарний номер",
    "Заводський номер", "Підрозділ бригади", "Підрозділ батальйону",
    "Примітка",
]
REGISTRY_HEADERS = [
    "№ накладної", "Дата накладної", "Спосіб надходження", "Постачальник",
    "Одержувач", "№ позиції", "Номенклатурний код", "Найменування майна",
    "Од. виміру", "Тип майна", "Кількість", "Ціна", "Сума",
    "Інвентарний номер", "Заводський номер", "Підрозділ бригади",
    "Підрозділ батальйону", "Примітка", "ID транзакції",
]


@dataclass(frozen=True)
class InvoiceItem:
    line: int
    code: str
    name: str
    unit: str
    asset_type: str
    quantity: Decimal
    price: Decimal
    inventory_number: str
    serial_number: str
    brigade: str
    battalion: str
    note: str


@dataclass(frozen=True)
class Invoice:
    number: str
    invoice_date: date
    arrival_method: str
    supplier: str
    recipient: str
    items: tuple[InvoiceItem, ...]


def _number(value: object, label: str, row: int, *, positive: bool) -> Decimal:
    try:
        number = Decimal(str(value).strip().replace(" ", "").replace(",", "."))
    except (InvalidOperation, ValueError):
        raise ValueError(f"Рядок {row}: {label} має бути числом.") from None
    if not number.is_finite() or (number <= 0 if positive else number < 0):
        raise ValueError(f"Рядок {row}: некоректне значення «{label}».")
    return number


def _date(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for pattern in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(value).strip(), pattern).date()
        except ValueError:
            pass
    raise ValueError("Вкажіть дату накладної у форматі дд.мм.рррр.")


def create_invoice_template(path: Path) -> None:
    """Створює редагований Excel-шаблон, який програма може зчитати назад."""
    wb = Workbook()
    ws = wb.active
    ws.title = TEMPLATE_SHEET
    ws.merge_cells("A1:L1")
    ws["A1"] = "НАКЛАДНА НА МАЙНО"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="24493D")
    ws["A1"].alignment = Alignment(horizontal="center")
    for cell, label in (("A2", "№ накладної"), ("D2", "Дата накладної"),
                        ("A3", "Спосіб надходження"), ("D3", "Постачальник"),
                        ("A4", "Одержувач")):
        ws[cell] = label
        ws[cell].font = Font(bold=True)
    ws["E2"].number_format = "DD.MM.YYYY"
    for column, header in enumerate(ITEM_HEADERS, 1):
        cell = ws.cell(6, column, header)
        cell.fill = PatternFill("solid", fgColor="24493D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(column)].width = (
            38 if column == 3 else 24 if column in (2, 8, 9, 10, 11, 12) else 16
        )
    ws.row_dimensions[6].height = 36
    ws.freeze_panes = "C7"
    ws.print_options.horizontalCentered = True
    ws.sheet_view.showGridLines = False
    wb.save(path)


def read_invoice(path: Path) -> Invoice:
    """Читає тільки наш шаблон; перевіряє всі рядки до зміни облікової книги."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if TEMPLATE_SHEET not in wb.sheetnames:
            raise ValueError("У файлі немає аркуша «Накладна» із шаблону Oblik.")
        ws = wb[TEMPLATE_SHEET]
        headers = [ws.cell(6, col).value for col in range(1, len(ITEM_HEADERS) + 1)]
        if headers != ITEM_HEADERS:
            raise ValueError("Структуру колонок накладної змінено. Скористайтеся шаблоном Oblik.")
        number = str(ws["B2"].value or "").strip()
        if not number:
            raise ValueError("Вкажіть номер накладної в B2.")
        invoice_date = _date(ws["E2"].value)
        arrival_method = str(ws["B3"].value or "").strip()
        if not arrival_method:
            raise ValueError("Вкажіть спосіб надходження в B3.")
        supplier = str(ws["E3"].value or "").strip()
        recipient = str(ws["B4"].value or "").strip()
        items = []
        for row in range(7, ws.max_row + 1):
            cells = [ws.cell(row, col).value for col in range(1, 13)]
            if all(value is None or str(value).strip() == "" for value in cells):
                continue
            text = lambda col: str(cells[col] or "").strip()
            if not text(2):
                raise ValueError(f"Рядок {row}: вкажіть найменування майна.")
            quantity = _number(cells[5], "Кількість", row, positive=True)
            price = _number(cells[6], "Ціна", row, positive=False)
            items.append(InvoiceItem(
                line=len(items) + 1, code=text(1), name=text(2), unit=text(3),
                asset_type=text(4) or "Запаси", quantity=quantity, price=price,
                inventory_number=text(7), serial_number=text(8),
                brigade=text(9), battalion=text(10), note=text(11),
            ))
        if not items:
            raise ValueError("Додайте хоча б одну позицію, починаючи з рядка 7.")
        return Invoice(number, invoice_date, arrival_method, supplier,
                       recipient, tuple(items))
    finally:
        wb.close()
