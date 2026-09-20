import sys
from pathlib import Path
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.main import OblikWorkbook, MAIN_HEADERS, SHEET_CURRENT, SHEET_MOVEMENT, SHEET_CHANGES, display_value
import pandas as pd


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
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "test.xlsx"
        model = OblikWorkbook()
        model.create_new(path)

        r1 = record("INV-001", "SN-001", "Н-1", "01.01.2024", "А-1", "02.01.2024", "РРЕБ")
        r2 = record("INV-001", "SN-001", "Н-2", "01.02.2024", "А-2", "02.02.2024", "1 МБ")
        row1 = model.append_record(SHEET_MOVEMENT, r1, "test")
        ws_move = model.wb[SHEET_MOVEMENT]
        assert row1 == 2
        assert ws_move[f"A{row1}"].value == "=ROW()-1"
        assert ws_move[f"R{row1}"].value == f"=P{row1}*Q{row1}"
        assert display_value(pd.NaT) == ""

        # Як і в інтерфейсі: перевіряємо майбутній запис ДО його збереження.
        # Те саме майно з іншими документами не має бути повним дублем.
        dup = model.score_candidate_duplicate(r2)
        assert dup.score < 5
        row2 = model.append_record(SHEET_MOVEMENT, r2, "test")
        assert ws_move[f"A{row2}"].value == "=ROW()-1"
        assert ws_move[f"R{row2}"].value == f"=P{row2}*Q{row2}"

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
        assert len(model.dataframe(SHEET_CHANGES)) >= 3

        model.save()
        assert path.exists() and path.stat().st_size > 0


if __name__ == "__main__":
    main()
    print("SMOKE TEST OK")
