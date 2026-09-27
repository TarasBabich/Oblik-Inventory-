import sys
from pathlib import Path
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.main import (OblikWorkbook, MAIN_HEADERS, SHEET_CURRENT, SHEET_MOVEMENT, SHEET_CHANGES, display_value, calculate_total, calculate_unit_price, AppSettingsStore, format_inventory_number, next_inventory_number, TRANSACTION_ID_HEADER, OPERATION_TYPE_HEADER, generate_transaction_id, format_transaction_id, parse_transaction_sequence, SUMMARY_GROUP_FIELDS, SUMMARY_OUTPUT_HEADERS, SHEET_SUMMARY, SHEET_STAFF, STAFF_HEADERS, staff_ui_numbered_values, estimate_staff_row_height, STAFF_MIN_COLUMN_WIDTH, STAFF_MAX_COLUMN_WIDTH, STAFF_MIN_ROW_HEIGHT, STAFF_MAX_ROW_HEIGHT)
import pandas as pd
import flet as ft
import flet_datatable2 as fdt
from types import SimpleNamespace
from src.main import FletOblikApp


def record(inv, serial, order, order_date, act, act_date, location, qty=1):
    data = {h: None for h in MAIN_HEADERS}
    data[MAIN_HEADERS[1]] = order
    data[MAIN_HEADERS[2]] = order_date
    data[MAIN_HEADERS[3]] = act
    data[MAIN_HEADERS[4]] = act_date
    data[MAIN_HEADERS[8]] = "Необоротний актив"
    data[MAIN_HEADERS[9]] = inv
    data[MAIN_HEADERS[12]] = "Тестовий засіб"
    data[MAIN_HEADERS[15]] = 1250.50
    data[MAIN_HEADERS[16]] = qty
    data[MAIN_HEADERS[18]] = serial
    data[MAIN_HEADERS[19]] = location
    data[MAIN_HEADERS[22]] = act_date
    return data


def main():
    # Flet 1.x removed deprecated aliases like WHITE24/WHITE70.
    assert ft.Colors.WHITE_24
    assert ft.Colors.WHITE_70
    assert calculate_total(1250.50, 4) == 5002.0
    unit_price = calculate_unit_price(100, 3)
    assert unit_price is not None
    assert abs(unit_price * 3 - 100) < 1e-10
    assert calculate_unit_price(100, 0) is None
    assert generate_transaction_id(1) == "TX-001"
    assert generate_transaction_id(2) == "TX-002"
    assert generate_transaction_id(999) == "TX-999"
    assert generate_transaction_id(1000) == "TX-001000"
    assert generate_transaction_id(1001) == "TX-001001"
    assert format_transaction_id(1000000) == "TX-001000000"
    assert parse_transaction_sequence("TX-001001") == 1001
    assert parse_transaction_sequence("TX-ABC") is None

    scrollbar = ft.Scrollbar(
        orientation=ft.ScrollbarOrientation.BOTTOM,
        thumb_visibility=True,
        track_visibility=True,
        interactive=True,
        thickness=12,
        radius=8,
    )
    assert scrollbar.orientation == ft.ScrollbarOrientation.BOTTOM
    assert scrollbar.thumb_visibility is True

    sticky_table = fdt.DataTable2(
        columns=[ft.DataColumn(label=ft.Text("ID"))],
        rows=[],
        fixed_top_rows=1,
        fixed_left_columns=1,
        visible_horizontal_scroll_bar=True,
        visible_vertical_scroll_bar=True,
        min_width=500,
    )
    assert sticky_table.fixed_top_rows == 1
    assert sticky_table.fixed_left_columns == 1
    assert sticky_table.visible_horizontal_scroll_bar is True

    hover_container = ft.Container(on_hover=lambda e: None)
    row2 = fdt.DataRow2(
        cells=[ft.DataCell(ft.Text("test"))],
        on_tap=lambda e: None,
    )
    popup = ft.PopupMenuButton(
        on_open=lambda e: None,
        items=[ft.PopupMenuItem(content="Дія")],
    )
    hidden_action_host = ft.Container(
        visible=True,
        opacity=0.0,
        content=ft.PopupMenuButton(
            disabled=True,
            items=[ft.PopupMenuItem(content="Дія")],
        ),
    )
    cell = ft.DataCell(hidden_action_host)
    assert hover_container.on_hover is not None
    assert row2.on_tap is not None
    assert popup.on_open is not None
    assert cell.content.visible is True
    assert cell.content.opacity == 0.0
    assert cell.content.content.disabled is True

    summary_dropdown = ft.Dropdown(
        value="Узагальнена назва",
        options=[
            ft.DropdownOption(key=label, text=label)
            for label in SUMMARY_GROUP_FIELDS
        ],
        on_select=lambda e: None,
    )
    assert summary_dropdown.value == "Узагальнена назва"

    # У Flet-відображенні «Штат» формули нумерації не показуються як текст.
    staff_values = [
        "=ROW()-1", "NOM-A", "Назва А", "од.", 5, None,
        "=ROW()-1", "NOM-A", "Назва А", "од.", 2, "РРЕБ", None,
    ]
    numbered, left_no, right_no, left_has, right_has = staff_ui_numbered_values(
        staff_values, 0, 0
    )
    assert left_has is True and right_has is True
    assert numbered[0] == 1
    assert numbered[6] == 1
    assert left_no == 1 and right_no == 1

    right_only = [
        "=ROW()-1", None, None, None, None, None,
        "=ROW()-1", "NOM-B", "Назва Б", "од.", 1, "1 МБ", None,
    ]
    numbered2, left_no, right_no, left_has, right_has = staff_ui_numbered_values(
        right_only, left_no, right_no
    )
    assert left_has is False and right_has is True
    assert numbered2[0] is None
    assert numbered2[6] == 2
    assert left_no == 1 and right_no == 2

    # Довге найменування повинно збільшувати висоту рядка, а не обрізатися.
    short_height = estimate_staff_row_height([
        1, "NOM-A", "Коротка назва", "од.", 1, None,
        1, "NOM-A", "Коротка назва", "од.", 1, "РРЕБ", None,
    ])
    long_height = estimate_staff_row_height([
        1, "NOM-A",
        "КОМПЛЕКС РАДІОЕЛЕКТРОННОЇ БОРОТЬБИ З БЕЗПІЛОТНИМИ ЛІТАЛЬНИМИ АПАРАТАМИ БУКОВЕЛЬ-АД",
        "од.", 1, None,
        1, "NOM-A",
        "КОМПЛЕКС РАДІОЕЛЕКТРОННОЇ БОРОТЬБИ З БЕЗПІЛОТНИМИ ЛІТАЛЬНИМИ АПАРАТАМИ БУКОВЕЛЬ-АД",
        "од.", 1, "РРЕБ", None,
    ])
    assert short_height >= 70
    assert long_height > short_height

    # Excel-like resize використовує GestureDetector та відповідні курсори.
    column_resize = ft.GestureDetector(
        width=8,
        mouse_cursor=ft.MouseCursor.RESIZE_LEFT_RIGHT,
        on_horizontal_drag_update=lambda e: None,
    )
    row_resize = ft.GestureDetector(
        height=8,
        mouse_cursor=ft.MouseCursor.RESIZE_UP_DOWN,
        on_vertical_drag_update=lambda e: None,
    )
    assert column_resize.mouse_cursor == ft.MouseCursor.RESIZE_LEFT_RIGHT
    assert row_resize.mouse_cursor == ft.MouseCursor.RESIZE_UP_DOWN
    # Flet надсилає загальний рух від початку жесту і окремий приріст.
    # Ширина/висота повинні змінюватися саме на приріст кожної події.
    drag = SimpleNamespace(primary_delta=4, local_delta=SimpleNamespace(x=80, y=80))
    assert FletOblikApp._drag_delta(drag, "x") == 4
    assert FletOblikApp._drag_delta(drag, "y") == 4
    # Послідовні позиції курсора змінюють ширину точно на пройдені пікселі,
    # а в книзі зберігається кінцева ширина цієї колонки.
    drag_app = object.__new__(FletOblikApp)
    drag_app.model = OblikWorkbook()
    drag_app.model.wb = __import__("openpyxl").Workbook()
    drag_app.model.wb.active.title = SHEET_STAFF
    drag_app.staff_column_widths = {1: 100.0}
    drag_app.staff_column_controls = {1: [ft.Container(width=100)]}
    drag_app.staff_grid_control = None
    drag_app.staff_column_drag_x = {}
    drag_app.page = SimpleNamespace(update=lambda: None)
    drag_app.status = ft.Text("")
    pos = lambda x: SimpleNamespace(global_position=SimpleNamespace(x=x))
    drag_app._start_staff_column_drag(1, pos(200))
    drag_app._pan_staff_column(1, pos(212))
    drag_app._pan_staff_column(1, pos(225))
    assert drag_app.staff_column_widths[1] == 125
    assert drag_app.staff_column_controls[1][0].width == 125
    assert drag_app.model.wb[SHEET_STAFF].column_dimensions["A"].width == 15.62
    drag_app._end_staff_column_drag(1)
    assert 1 not in drag_app.staff_column_drag_x
    assert STAFF_MIN_COLUMN_WIDTH < STAFF_MAX_COLUMN_WIDTH
    assert STAFF_MIN_ROW_HEIGHT < STAFF_MAX_ROW_HEIGHT

    assert format_inventory_number("ОВТ-", 25, 6) == "ОВТ-000025"
    assert format_inventory_number("100-", 7, 4, "/26") == "100-0007/26"
    inv_settings = {
        "inventory_prefix": "ОВТ-",
        "inventory_suffix": "",
        "inventory_next_number": "1",
        "inventory_digits": "4",
    }
    candidate, counter = next_inventory_number(
        inv_settings,
        ["ОВТ-0001", "ОВТ-0002", "ОВТ-0004"],
    )
    assert candidate == "ОВТ-0003"
    assert counter == 3

    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "test.xlsx"
        settings = AppSettingsStore()
        settings.path = Path(tmp) / "oblik_settings.json"
        saved_path = settings.save({
            "unit_number": "А1234",
            "commander_rank": "полковник",
            "commander_name": "Тестовий Командир",
            "document_start_number": "25",
            "inventory_prefix": "ОВТ-",
            "inventory_suffix": "/26",
            "inventory_next_number": "15",
            "inventory_digits": "5",
            "index_coefficient_2023": "1.10",
            "index_coefficient_2024": "1.25",
            "index_coefficient_2025": "1.40",
            "commission_chair_position": "Начальник служби",
            "commission_chair_rank": "майор",
            "commission_chair_name": "Голова Комісії",
            "commission_members": [
                {"position": "Офіцер", "rank": "капітан", "name": "Член Один"},
                {"position": "Інженер", "rank": "старший лейтенант", "name": "Член Два"},
            ],
        })
        loaded_settings = settings.load()
        assert saved_path.exists()
        assert loaded_settings["unit_number"] == "А1234"
        assert loaded_settings["commander_rank"] == "полковник"
        assert loaded_settings["commander_name"] == "Тестовий Командир"
        assert loaded_settings["document_start_number"] == "25"
        assert loaded_settings["inventory_prefix"] == "ОВТ-"
        assert loaded_settings["inventory_suffix"] == "/26"
        assert loaded_settings["inventory_next_number"] == "15"
        assert loaded_settings["inventory_digits"] == "5"
        assert loaded_settings["index_coefficient_2023"] == "1.10"
        assert loaded_settings["index_coefficient_2024"] == "1.25"
        assert loaded_settings["index_coefficient_2025"] == "1.40"
        assert loaded_settings["commission_chair_name"] == "Голова Комісії"
        assert len(loaded_settings["commission_members"]) == 2
        assert loaded_settings["commission_members"][1]["name"] == "Член Два"

        model = OblikWorkbook()
        model.create_new(path)

        r1 = record("INV-001", "SN-001", "Н-1", "01.01.2024", "А-1", "02.01.2024", "РРЕБ")
        r2 = record("INV-001", "SN-001", "Н-2", "01.02.2024", "А-2", "02.02.2024", "1 МБ")
        for item in (r1, r2):
            item[MAIN_HEADERS[7]] = "NOM-A"
            item[MAIN_HEADERS[10]] = "Номенклатура А"
            item[MAIN_HEADERS[11]] = "Засоби РЕБ"
            item[MAIN_HEADERS[13]] = "SAP-001"
            item[MAIN_HEADERS[23]] = "Справний"
        row1 = model.append_record(SHEET_MOVEMENT, r1, "test")
        ws_move = model.wb[SHEET_MOVEMENT]
        assert row1 == 2
        assert ws_move[f"A{row1}"].value == "=ROW()-1"
        assert ws_move[f"R{row1}"].value == f"=P{row1}*Q{row1}"
        tx_col = model.headers(SHEET_MOVEMENT).index(TRANSACTION_ID_HEADER) + 1
        tx1 = ws_move.cell(row1, tx_col).value
        assert tx1 == "TX-001"
        op_col = model.headers(SHEET_MOVEMENT).index(OPERATION_TYPE_HEADER) + 1
        assert ws_move.cell(row1, op_col).value == "Первинний запис"
        assert display_value(pd.NaT) == ""

        # Як і в інтерфейсі: перевіряємо майбутній запис ДО його збереження.
        # Те саме майно з іншими документами не має бути повним дублем.
        dup = model.score_candidate_duplicate(r2)
        assert dup.score < 5
        r2[OPERATION_TYPE_HEADER] = "Переміщення"
        row2 = model.append_record(SHEET_MOVEMENT, r2, "test")
        assert ws_move[f"A{row2}"].value == "=ROW()-1"
        assert ws_move[f"R{row2}"].value == f"=P{row2}*Q{row2}"
        tx2 = ws_move.cell(row2, tx_col).value
        assert tx2 == "TX-002"
        assert ws_move.cell(row2, op_col).value == "Переміщення"

        # Службові транзакції руху не є дублями первинного оприбуткування.
        dup_map = model.duplicate_map()
        assert dup_map.get(row1) is None or dup_map[row1].score < 5
        assert row2 not in dup_map

        count, skipped = model.rebuild_current_state()
        assert count == 1
        current = model.dataframe(SHEET_CURRENT)
        assert len(current) == 1
        assert str(current.iloc[0][MAIN_HEADERS[19]]) == "1 МБ"
        assert model.wb[SHEET_CURRENT]["A2"].value == "=ROW()-1"
        assert model.wb[SHEET_CURRENT]["R2"].value == "=P2*Q2"

        movement = model.dataframe(SHEET_MOVEMENT)
        first_row = int(movement.iloc[0]["_excel_row"])
        edited = dict(r1)
        edited[MAIN_HEADERS[18]] = "SN-001-CORRECTED"
        changes = model.update_record(SHEET_MOVEMENT, first_row, edited, "уточнення")
        assert any(field == MAIN_HEADERS[18] for field, _, _ in changes)
        assert ws_move.cell(first_row, tx_col).value == tx1

        audit = model.dataframe(SHEET_CHANGES)
        assert len(audit) >= 3
        assert TRANSACTION_ID_HEADER in audit.columns
        assert tx1 in set(audit[TRANSACTION_ID_HEADER].dropna().astype(str))

        # Симулюємо старий UUID-ID: при відкритті він мігрує у наступний
        # послідовний номер і синхронізується з журналом.
        legacy_id = "TX-ABCDEF0123456789ABCDEF0123456789"
        ws_move.cell(row2, tx_col).value = legacy_id
        changes_ws = model.wb[SHEET_CHANGES]
        change_headers = model.headers(SHEET_CHANGES)
        audit_tx_col = change_headers.index(TRANSACTION_ID_HEADER) + 1
        audit_row = changes_ws.max_row + 1
        changes_ws.cell(audit_row, audit_tx_col, legacy_id)
        model.save()

        reloaded = OblikWorkbook()
        reloaded.load(path)
        reloaded_ws = reloaded.wb[SHEET_MOVEMENT]
        restored_tx2 = reloaded_ws.cell(row2, tx_col).value
        assert restored_tx2 == "TX-003"
        assert reloaded_ws[f"A{row2}"].value == "=ROW()-1"
        assert reloaded_ws[f"R{row2}"].value == f"=P{row2}*Q{row2}"
        reloaded_changes = reloaded.wb[SHEET_CHANGES]
        assert reloaded_changes.cell(audit_row, audit_tx_col).value == "TX-003"

        # Видалений номер не перевикористовується, бо лишається в Контролі змін.
        reloaded.delete_record(SHEET_MOVEMENT, row2, "test delete")
        row3 = reloaded.append_record(SHEET_MOVEMENT, r2, "after delete")
        assert reloaded.wb[SHEET_MOVEMENT].cell(row3, tx_col).value == "TX-004"

        # Перевіряємо три незалежні режими групування Зведеного.
        r3 = record("INV-002", "SN-002", "Н-3", "01.03.2024", "А-3", "02.03.2024", "Склад", qty=2)
        r3[MAIN_HEADERS[7]] = "NOM-B"
        r3[MAIN_HEADERS[10]] = "Номенклатура Б"
        r3[MAIN_HEADERS[11]] = "Засоби РЕБ"
        r3[MAIN_HEADERS[13]] = "SAP-001"
        r3[MAIN_HEADERS[23]] = "Несправний"
        r3["Примітка"] = "Потребує ремонту"

        # Штат: дві окремі таблиці A:E та G:M.
        ws_staff = reloaded.wb[SHEET_STAFF]
        assert [ws_staff.cell(1, c).value or "" for c in range(1, 14)] == STAFF_HEADERS
        assert ws_staff["D1"].value == "Од. виміру"
        assert ws_staff["E1"].value == "Узагальнена кількість за військову частину"
        assert ws_staff["F1"].value in (None, "")
        assert ws_staff["J1"].value == "Од. виміру"
        assert ws_staff["K1"].value == "Кількість"
        assert ws_staff["L1"].value == "Окремий підрозділ бригади"
        assert ws_staff["M1"].value == "Окремий підрозділ батальйону/дивізіону"

        ws_staff["B2"] = "NOM-A"
        ws_staff["C2"] = "Номенклатура А"
        ws_staff["D2"] = "од."
        ws_staff["E2"] = 5
        ws_staff["B3"] = "NOM-B"
        ws_staff["C3"] = "Номенклатура Б"
        ws_staff["D3"] = "од."
        ws_staff["E3"] = 3

        # Деталізація праворуч не повинна впливати на загальну штатну кількість.
        ws_staff["G2"] = 1
        ws_staff["H2"] = "NOM-A"
        ws_staff["I2"] = "Номенклатура А"
        ws_staff["J2"] = "од."
        ws_staff["K2"] = 2
        ws_staff["L2"] = "РРЕБ"
        ws_staff["M2"] = ""
        reloaded.append_record(SHEET_MOVEMENT, r3, "summary test")
        count, skipped = reloaded.rebuild_current_state()
        assert count == 2

        by_general = reloaded.build_summary_dataframe(MAIN_HEADERS[11])
        assert len(by_general) == 1
        general = by_general.iloc[0]
        assert general["Найменування"] == "Засоби РЕБ"
        assert general["Штат"] == 8
        assert general["За обліком"] == 3
        assert general["Наявні (справні)"] == 1
        assert general["Несправні"] == 2
        assert general["БПВ"] == 0
        assert abs(general["% забезпеченості справних"] - 12.5) < 1e-9
        assert "Потребує ремонту" in general["Примітка"]

        # Поле «Штатна потреба» поза аркушем «Штат» не повинно впливати
        # на зведення, навіть якщо в ньому є значення.
        orphan = record("INV-003", "SN-003", "Н-4", "01.04.2024", "А-4", "02.04.2024", "Склад")
        orphan[MAIN_HEADERS[7]] = "NOM-C"
        orphan[MAIN_HEADERS[10]] = "Номенклатура В"
        orphan[MAIN_HEADERS[11]] = "Без штатної відповідності"
        orphan[MAIN_HEADERS[13]] = "SAP-999"
        orphan[MAIN_HEADERS[14]] = 99
        orphan[MAIN_HEADERS[23]] = "Справний"
        reloaded.append_record(SHEET_MOVEMENT, orphan, "staff source test")
        reloaded.rebuild_current_state()
        orphan_summary = reloaded.build_summary_dataframe(MAIN_HEADERS[11])
        orphan_row = orphan_summary[
            orphan_summary["Найменування"] == "Без штатної відповідності"
        ].iloc[0]
        assert orphan_row["Штат"] == 0
        assert orphan_row["Наявні (справні)"] == 1
        assert pd.isna(orphan_row["% забезпеченості справних"])

        # Штатна позиція без жодної фактичної одиниці все одно має бути
        # присутня у Зведеному з нульовим поточним станом.
        ws_staff["B4"] = "NOM-D"
        ws_staff["C4"] = "Номенклатура Г"
        ws_staff["D4"] = "од."
        ws_staff["E4"] = 7

        by_nomenclature = reloaded.build_summary_dataframe(MAIN_HEADERS[10])
        assert len(by_nomenclature) == 4
        assert set(by_nomenclature["Найменування"]) == {
            "Номенклатура А",
            "Номенклатура Б",
            "Номенклатура В",
            "Номенклатура Г",
        }
        nomenclature_a = by_nomenclature[
            by_nomenclature["Найменування"] == "Номенклатура А"
        ].iloc[0]
        assert nomenclature_a["Штат"] == 5
        assert nomenclature_a["За обліком"] == 1
        assert nomenclature_a["Наявні (справні)"] == 1
        assert abs(nomenclature_a["% забезпеченості справних"] - 20.0) < 1e-9

        staff_only = by_nomenclature[
            by_nomenclature["Найменування"] == "Номенклатура Г"
        ].iloc[0]
        assert staff_only["Штат"] == 7
        assert staff_only["За обліком"] == 0
        assert staff_only["Наявні (справні)"] == 0
        assert staff_only["Несправні"] == 0
        assert "Немає в поточному стані" in staff_only["Примітка"]
        assert staff_only["% забезпеченості справних"] == 0

        by_sap = reloaded.rebuild_summary(MAIN_HEADERS[13])
        assert len(by_sap) == 3
        sap_001 = by_sap[by_sap["Найменування"] == "SAP-001"].iloc[0]
        assert sap_001["Штат"] == 8
        ws_summary = reloaded.wb[SHEET_SUMMARY]
        assert ws_summary["A1"].value == "№ з/п"
        assert ws_summary["B1"].value == "Номер матеріалу в SAP"
        assert ws_summary["C1"].value == "Штат"
        assert ws_summary["D1"].value == "За обліком"
        assert ws_summary["E1"].value == "Наявні (справні)"
        assert ws_summary["F1"].value == "Несправні"
        assert ws_summary["G1"].value == "БПВ"
        assert ws_summary["H1"].value == "% забезпеченості справних"
        assert ws_summary["I1"].value == "Примітка"
        # SAP-001 має штат 8 виключно з аркуша «Штат».
        summary_rows = {
            ws_summary.cell(row, 2).value: row
            for row in range(2, ws_summary.max_row + 1)
        }
        sap_row = summary_rows["SAP-001"]
        assert ws_summary.cell(sap_row, 3).value == 8
        assert ws_summary.cell(sap_row, 4).value == 3
        assert ws_summary.cell(sap_row, 5).value == 1
        assert ws_summary.cell(sap_row, 6).value == 2
        assert abs(ws_summary.cell(sap_row, 8).value - 12.5) < 1e-9

        # SAP-999 не має відповідності на аркуші «Штат»: штат = 0,
        # попри «Штатна потреба» = 99 у записі руху.
        orphan_sap_row = summary_rows["SAP-999"]
        assert ws_summary.cell(orphan_sap_row, 3).value == 0
        assert ws_summary.cell(orphan_sap_row, 5).value == 1
        assert ws_summary.cell(orphan_sap_row, 8).value is None

        reloaded.save()
        assert path.exists() and path.stat().st_size > 0


if __name__ == "__main__":
    main()
    print("SMOKE TEST OK")
