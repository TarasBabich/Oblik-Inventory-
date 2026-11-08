"""Наскрізна перевірка: Excel-шаблон → реєстр → рух майна → повторне відкриття."""

from pathlib import Path
from tempfile import TemporaryDirectory
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import load_workbook

from src.invoices import INVOICE_SHEET, create_invoice_template
from src.main import MAIN_HEADERS, OPERATION_TYPE_HEADER, SHEET_CURRENT, SHEET_MOVEMENT, OblikWorkbook


def main():
    with TemporaryDirectory() as directory:
        root = Path(directory)
        book_path = root / "Облік.xlsx"
        invoice_path = root / "Накладна.xlsx"
        model = OblikWorkbook()
        model.create_new(book_path)
        create_invoice_template(invoice_path)
        invoice = load_workbook(invoice_path)
        ws = invoice["Накладна"]
        ws["B2"] = "Н-42"
        ws["E2"] = "08.11.2026"
        ws["B3"] = "Передача від іншої частини"
        ws["E3"] = "Умовний постачальник"
        ws["B4"] = "Склад"
        ws["C7"] = "Прилад вимірювальний"
        ws["E7"] = "Необоротний актив"
        ws["F7"] = 1
        ws["G7"] = "1250,50"
        ws["H7"] = "ОВТ-0001"
        ws["C8"] = "Запасний кабель"
        ws["E8"] = "Запаси"
        ws["F8"] = 3
        ws["G8"] = 20
        invoice.save(invoice_path)

        assert model.import_invoice(invoice_path) == 2
        registry = model.dataframe(INVOICE_SHEET)
        movement = model.dataframe(SHEET_MOVEMENT)
        assert len(registry) == len(movement) == 2
        assert registry.iloc[0]["Спосіб надходження"] == "Передача від іншої частини"
        assert registry.iloc[0]["ID транзакції"] == movement.iloc[0]["ID транзакції"]
        assert movement.iloc[1][MAIN_HEADERS[9]] is None  # запас без інвентарного №
        assert movement.iloc[1][OPERATION_TYPE_HEADER] == "Оприбуткування за накладною"
        assert len(model.dataframe(SHEET_CURRENT)) == 1
        try:
            model.import_invoice(invoice_path)
            raise AssertionError("Повторне завантаження повинно відхилятися")
        except ValueError as error:
            assert "вже внесено" in str(error)

        model.save()
        reopened = OblikWorkbook()
        reopened.load(book_path)
        assert len(reopened.dataframe(INVOICE_SHEET)) == 2
        assert len(reopened.dataframe(SHEET_MOVEMENT)) == 2

        ws["F8"] = -1
        invoice.save(invoice_path)
        try:
            reopened.import_invoice(invoice_path)
            raise AssertionError("Некоректну позицію потрібно відхилити")
        except ValueError:
            assert len(reopened.dataframe(INVOICE_SHEET)) == 2


if __name__ == "__main__":
    main()
    print("INVOICE TEST OK")
