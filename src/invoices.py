# -*- coding: utf-8 -*-
"""Накладні (вимоги) за додатком 25 до Інструкції МОУ № 440."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

SHEET_INVOICE_INTERNAL = "Накладна на підрозділи"
SHEET_INVOICE_EXTERNAL = "Накладна на іншу вч"
SHEET_TRANSFER_REGISTER = "Список переміщень"

INVOICE_INTERNAL = "Переміщення між підрозділами військової частини"
INVOICE_EXTERNAL = "Передача між військовими частинами"

TRANSFER_HEADERS = [
    "№ з/п", "Найменування майна", "Заводський номер майна",
    "Код номенклатури", "Од. виміру", "Категорія", "Ціна",
    "Кількість", "Сума",
    "З якого окремого підрозділу бригади переміщено",
    "З якого підрозділу батальйону/дивізіону переміщено",
    "До якого окремого підрозділу бригади переміщено",
    "До якого підрозділу батальйону/дивізіону переміщено",
    "Військова частина-відправник", "Код ЄДРПОУ відправника",
    "Військова частина-одержувач", "Код ЄДРПОУ одержувача",
    "Номер накладної", "Дата накладної", "Тип переміщення",
    "Підстава (мета)", "Передав", "Прийняв",
    "Шлях до файлу накладної", "Дата і час створення запису",
]


@dataclass
class InvoiceItem:
    name: str
    serial_number: str = ""
    nomenclature_code: str = ""
    unit: str = ""
    category: str = ""
    price: float = 0.0
    quantity_sent: float = 0.0
    quantity_received: float = 0.0
    note: str = ""

    @property
    def quantity(self) -> float:
        return self.quantity_received or self.quantity_sent

    @property
    def total(self) -> float:
        return round(float(self.price or 0) * float(self.quantity or 0), 2)


@dataclass
class InvoiceData:
    invoice_type: str = INVOICE_INTERNAL
    number: str = ""
    issue_date: date = field(default_factory=date.today)
    operation_date: date = field(default_factory=date.today)
    valid_until: Optional[date] = None
    place: str = ""
    service_name: str = ""
    basis: str = ""
    responsible_receiver: str = ""
    sender: str = ""
    receiver: str = ""
    from_brigade_unit: str = ""
    from_battalion_unit: str = ""
    to_brigade_unit: str = ""
    to_battalion_unit: str = ""
    sender_unit_name: str = ""
    sender_edrpou: str = ""
    receiver_unit_name: str = ""
    receiver_edrpou: str = ""
    leader_position: str = "Керівник (посадова особа, начальник служби)"
    leader_name: str = ""
    items: list[InvoiceItem] = field(default_factory=list)


def _thin_border() -> Border:
    side = Side(style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)


def _date_text(value: Optional[date]) -> str:
    return value.strftime("%d.%m.%Y") if value else ""


def _parse_date(value: Any) -> Optional[date]:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Невірна дата: {text}. Використовуйте формат дд.мм.рррр.")


def parse_invoice_date(value: Any, default: Optional[date] = None) -> date:
    parsed = _parse_date(value)
    return parsed or default or date.today()


def _ua_int(number: int) -> str:
    if number == 0:
        return "нуль"
    ones = ["", "один", "два", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять"]
    ones_f = ["", "одна", "дві", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять"]
    teens = ["десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", "п’ятнадцять", "шістнадцять", "сімнадцять", "вісімнадцять", "дев’ятнадцять"]
    tens = ["", "", "двадцять", "тридцять", "сорок", "п’ятдесят", "шістдесят", "сімдесят", "вісімдесят", "дев’яносто"]
    hundreds = ["", "сто", "двісті", "триста", "чотириста", "п’ятсот", "шістсот", "сімсот", "вісімсот", "дев’ятсот"]
    scales = [
        (1_000_000_000, ("мільярд", "мільярди", "мільярдів"), False),
        (1_000_000, ("мільйон", "мільйони", "мільйонів"), False),
        (1_000, ("тисяча", "тисячі", "тисяч"), True),
    ]

    def form(n: int, words: tuple[str, str, str]) -> str:
        last2 = n % 100
        last = n % 10
        if 11 <= last2 <= 14:
            return words[2]
        if last == 1:
            return words[0]
        if 2 <= last <= 4:
            return words[1]
        return words[2]

    def chunk(n: int, feminine: bool = False) -> str:
        parts: list[str] = []
        if n >= 100:
            parts.append(hundreds[n // 100])
            n %= 100
        if 10 <= n <= 19:
            parts.append(teens[n - 10])
        else:
            if n >= 20:
                parts.append(tens[n // 10])
                n %= 10
            if n:
                parts.append((ones_f if feminine else ones)[n])
        return " ".join(parts)

    parts: list[str] = []
    remainder = number
    for divisor, words, feminine in scales:
        amount, remainder = divmod(remainder, divisor)
        if amount:
            parts.extend([chunk(amount, feminine), form(amount, words)])
    if remainder:
        parts.append(chunk(remainder))
    return " ".join(parts)


def amount_in_words(value: float) -> str:
    amount = round(float(value or 0), 2)
    hryvnias = int(amount)
    kopiyky = int(round((amount - hryvnias) * 100))
    return f"{_ua_int(hryvnias)} грн {kopiyky:02d} коп."


def quantity_in_words(value: float) -> str:
    number = float(value or 0)
    if number.is_integer():
        return _ua_int(int(number))
    return str(number).replace(".", ",")


def _style_invoice_sheet(ws, data: InvoiceData, settings: dict[str, Any]) -> None:
    ws.title = "Накладна (вимога)"
    widths = [5, 29, 13, 10, 10, 13, 13, 13, 14, 20]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.sheet_view.showGridLines = False

    def merge(address: str, value: Any, size=10, bold=False, align="left"):
        ws.merge_cells(address)
        cell = ws[address.split(":")[0]]
        cell.value = value
        cell.font = Font(name="Times New Roman", size=size, bold=bold)
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)

    merge("G1:J1", "Додаток 25")
    merge("G2:J2", "до Інструкції з обліку військового майна")
    merge("G3:J3", "у Збройних Силах України")
    merge("G4:J4", "(пункт 24 розділу IV)")
    unit_name = data.sender_unit_name or settings.get("unit_name") or settings.get("unit_number") or ""
    merge("A6:F6", unit_name, 11, True, "center")
    merge("A7:F7", "(найменування юридичної особи)", 8, False, "center")
    merge("A8:C8", "Код згідно з ЄДРПОУ")
    merge("D8:F8", data.sender_edrpou or settings.get("unit_edrpou", ""), 10, True, "center")
    merge("A9:F9", f"Дійсна до {_date_text(data.valid_until)}")
    merge("A11:J11", f"НАКЛАДНА (ВИМОГА) № {data.number}", 14, True, "center")
    merge("A12:J12", data.place, 10, False, "center")
    merge("A13:J13", "(місце складання)", 8, False, "center")
    merge("A14:J14", _date_text(data.issue_date), 10, False, "center")
    merge("A15:J15", "(дата складання)", 8, False, "center")
    merge("A17:E17", f"Дата операції: {_date_text(data.operation_date)}")
    merge("F17:J17", f"Служба забезпечення: {data.service_name or settings.get('service_name', '')}")
    merge("A18:E18", f"Вид операції: {data.invoice_type}")
    merge("F18:J18", f"Підстава (мета): {data.basis}")
    merge("A19:J19", f"Відповідальний одержувач: {data.responsible_receiver}")
    merge("A20:E20", f"Передає: {data.sender}")
    merge("F20:J20", f"Приймає: {data.receiver}")

    headers = ["№ з/п", "Назва військового майна або однорідна група (вид)", "Код номенклатури", "Одиниця виміру", "Категорія (сорт)", "Вартість за одиницю", "Кількість", "", "Сума", "Примітка"]
    for col, value in enumerate(headers, 1):
        ws.cell(22, col, value)
    ws.merge_cells("G22:H22")
    for col in (1, 2, 3, 4, 5, 6, 9, 10):
        ws.merge_cells(start_row=22, start_column=col, end_row=23, end_column=col)
    ws["G23"] = "Відправлено\n(вимагається)"
    ws["H23"] = "Прийнято\n(відпущено)"
    for col in range(1, 11):
        ws.cell(24, col, str(col))
    for row in range(22, 25):
        for col in range(1, 11):
            cell = ws.cell(row, col)
            cell.font = Font(name="Times New Roman", size=9, bold=row < 24)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = PatternFill("solid", fgColor="E7E6E6") if row < 24 else PatternFill()
            cell.border = _thin_border()
    ws.row_dimensions[22].height = 34
    ws.row_dimensions[23].height = 30

    current = 25
    for index, item in enumerate(data.items, 1):
        top, bottom = current, current + 1
        for col in (1, 3, 4, 5, 6, 7, 8, 9, 10):
            ws.merge_cells(start_row=top, start_column=col, end_row=bottom, end_column=col)
        values = {
            1: index, 2: item.name, 3: item.nomenclature_code, 4: item.unit,
            5: item.category, 6: item.price, 7: item.quantity_sent,
            8: item.quantity_received, 9: item.total, 10: item.note,
        }
        for col, value in values.items():
            ws.cell(top, col, value)
        # Якщо заводського номера немає, нижній рядок залишається повністю порожнім.
        ws.cell(bottom, 2, item.serial_number.strip() if item.serial_number.strip() else None)
        for row in (top, bottom):
            for col in range(1, 11):
                cell = ws.cell(row, col)
                cell.font = Font(name="Times New Roman", size=9)
                cell.alignment = Alignment(horizontal="left" if col in (2, 10) else "center", vertical="center", wrap_text=True)
                cell.border = _thin_border()
        ws.cell(top, 6).number_format = "#,##0.00"
        ws.cell(top, 7).number_format = "#,##0.###"
        ws.cell(top, 8).number_format = "#,##0.###"
        ws.cell(top, 9).number_format = "#,##0.00"
        ws.row_dimensions[top].height = 24
        ws.row_dimensions[bottom].height = 20
        current += 2

    if not data.items:
        current = 27
        for col in (1, 3, 4, 5, 6, 7, 8, 9, 10):
            ws.merge_cells(start_row=25, start_column=col, end_row=26, end_column=col)
        for row in (25, 26):
            for col in range(1, 11):
                ws.cell(row, col).border = _thin_border()

    total_row = current
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=6)
    ws.cell(total_row, 1, "Всього")
    ws.cell(total_row, 7, sum(i.quantity_sent for i in data.items))
    ws.cell(total_row, 8, sum(i.quantity_received for i in data.items))
    ws.cell(total_row, 9, sum(i.total for i in data.items))
    for col in range(1, 11):
        cell = ws.cell(total_row, col)
        cell.font = Font(name="Times New Roman", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="E7E6E6")
        cell.border = _thin_border()
    ws.cell(total_row, 9).number_format = "#,##0.00"

    footer = total_row + 2
    total_qty = sum(i.quantity for i in data.items)
    total_sum = sum(i.total for i in data.items)
    footer_lines = [
        f"{data.leader_position}: __________________  {data.leader_name}",
        "                                     (підпис)                 (власне ім’я та прізвище)",
        "",
        f"Всього передано {quantity_in_words(total_qty)} одиниць,",
        "                         (кількість прописом)",
        f"на суму {amount_in_words(total_sum)}",
        "                         (сума прописом)",
        "",
        "Матеріально відповідальні особи:",
        f"здав:      {data.sender}     __________________",
        "                         (посада, власне ім’я та прізвище)       (підпис)",
        "",
        f"прийняв: {data.receiver}     __________________",
        "                         (посада, власне ім’я та прізвище)       (підпис)",
    ]
    for offset, line in enumerate(footer_lines):
        row = footer + offset
        merge(f"A{row}:J{row}", line, 10 if "(" not in line else 8, "Матеріально" in line, "left")

    last_row = footer + len(footer_lines) - 1
    ws.freeze_panes = "A25"
    ws.print_title_rows = "22:24"
    ws.print_area = f"A1:J{last_row}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.35, bottom=0.35, header=0.15, footer=0.15)
    ws.oddFooter.center.text = "Сторінка &P з &N"


def _style_transfer_register(ws) -> None:
    for col, header in enumerate(TRANSFER_HEADERS, 1):
        ws.cell(1, col, header)
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border()
    ws.row_dimensions[1].height = 54
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(TRANSFER_HEADERS))}1"
    widths = [8, 30, 18, 18, 12, 12, 14, 14, 14, 25, 25, 25, 25, 22, 18, 22, 18, 18, 16, 22, 24, 22, 22, 38, 20]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width


def ensure_invoice_sheets(wb, settings: Optional[dict[str, Any]] = None) -> bool:
    """Додає три аркуші до нової або старої книги без видалення наявних даних."""
    changed = False
    settings = settings or {}
    if SHEET_INVOICE_INTERNAL not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_INVOICE_INTERNAL)
        _style_invoice_sheet(ws, InvoiceData(invoice_type=INVOICE_INTERNAL), settings)
        changed = True
    if SHEET_INVOICE_EXTERNAL not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_INVOICE_EXTERNAL)
        _style_invoice_sheet(ws, InvoiceData(invoice_type=INVOICE_EXTERNAL), settings)
        changed = True
    if SHEET_TRANSFER_REGISTER not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_TRANSFER_REGISTER)
        _style_transfer_register(ws)
        changed = True
    return changed


def append_transfer_records(ws, data: InvoiceData, output_path: Path) -> int:
    """Записує кожну позицію накладної окремим рядком до реєстру."""
    if ws.max_row < 1 or ws.cell(1, 1).value != TRANSFER_HEADERS[0]:
        _style_transfer_register(ws)
    start = ws.max_row + 1
    created_at = datetime.now()
    for offset, item in enumerate(data.items):
        row = start + offset
        values = [
            row - 1, item.name, item.serial_number or None, item.nomenclature_code,
            item.unit, item.category, item.price, item.quantity, item.total,
            data.from_brigade_unit, data.from_battalion_unit,
            data.to_brigade_unit, data.to_battalion_unit,
            data.sender_unit_name, data.sender_edrpou,
            data.receiver_unit_name, data.receiver_edrpou,
            data.number, data.issue_date, data.invoice_type, data.basis,
            data.sender, data.receiver, str(output_path), created_at,
        ]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="hair", color="BFBFBF"))
        ws.cell(row, 7).number_format = "#,##0.00"
        ws.cell(row, 8).number_format = "#,##0.###"
        ws.cell(row, 9).number_format = "#,##0.00"
        ws.cell(row, 19).number_format = "dd.mm.yyyy"
        ws.cell(row, 25).number_format = "dd.mm.yyyy hh:mm"
    return len(data.items)


def generate_invoice_file(
    output_path: Path,
    data: InvoiceData,
    settings: Optional[dict[str, Any]] = None,
) -> Path:
    if not data.number.strip():
        raise ValueError("Вкажіть номер накладної.")
    if not data.items:
        raise ValueError("Додайте хоча б одну позицію майна.")
    if any(not item.name.strip() for item in data.items):
        raise ValueError("У кожній позиції має бути найменування майна.")
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")
    wb = Workbook()
    ws = wb.active
    _style_invoice_sheet(ws, data, settings or {})
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(output_path)
    return output_path
