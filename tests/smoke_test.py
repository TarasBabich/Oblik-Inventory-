import sys
from pathlib import Path
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.main import (OblikWorkbook, MAIN_HEADERS, SHEET_CURRENT, SHEET_MOVEMENT, SHEET_CHANGES, display_value, calculate_total, calculate_unit_price, AppSettingsStore, format_inventory_number, next_inventory_number, TRANSACTION_ID_HEADER, generate_transaction_id)
import pandas as pd
import flet as ft


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
    generated_tx = generate_transaction_id()
    assert generated_tx.startswith("TX-")
    assert len(generated_tx) == 35

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
        row1 = model.append_record(SHEET_MOVEMENT, r1, "test")
        ws_move = model.wb[SHEET_MOVEMENT]
        assert row1 == 2
        assert ws_move[f"A{row1}"].value == "=ROW()-1"
        assert ws_move[f"R{row1}"].value == f"=P{row1}*Q{row1}"
        tx_col = model.headers(SHEET_MOVEMENT).index(TRANSACTION_ID_HEADER) + 1
        tx1 = ws_move.cell(row1, tx_col).value
        assert str(tx1).startswith("TX-")
        assert display_value(pd.NaT) == ""

        # Як і в інтерфейсі: перевіряємо майбутній запис ДО його збереження.
        # Те саме майно з іншими документами не має бути повним дублем.
        dup = model.score_candidate_duplicate(r2)
        assert dup.score < 5
        row2 = model.append_record(SHEET_MOVEMENT, r2, "test")
        assert ws_move[f"A{row2}"].value == "=ROW()-1"
        assert ws_move[f"R{row2}"].value == f"=P{row2}*Q{row2}"
        tx2 = ws_move.cell(row2, tx_col).value
        assert str(tx2).startswith("TX-")
        assert tx2 != tx1

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

        # Симулюємо старий рядок без ID: при наступному відкритті він має
        # отримати ID автоматично, не змінюючи Excel-рядок чи формули.
        ws_move.cell(row2, tx_col).value = None
        model.save()
        reloaded = OblikWorkbook()
        reloaded.load(path)
        reloaded_ws = reloaded.wb[SHEET_MOVEMENT]
        restored_tx2 = reloaded_ws.cell(row2, tx_col).value
        assert str(restored_tx2).startswith("TX-")
        assert restored_tx2 != tx1
        assert reloaded_ws[f"A{row2}"].value == "=ROW()-1"
        assert reloaded_ws[f"R{row2}"].value == f"=P{row2}*Q{row2}"

        reloaded.save()
        assert path.exists() and path.stat().st_size > 0


if __name__ == "__main__":
    main()
    print("SMOKE TEST OK")
