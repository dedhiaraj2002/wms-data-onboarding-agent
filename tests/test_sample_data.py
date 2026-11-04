import json
import sqlite3

import pandas as pd
import pytest

from scripts.generate_sample_data import CUSTOMER, RULES, generate


def test_excel_contains_twenty_rows_and_intentional_defects(tmp_path):
    workbook, _ = generate(tmp_path)
    rows = pd.read_excel(workbook, sheet_name="locations")
    assert len(rows) == 20
    assert set(rows.customer_id) == {CUSTOMER}
    assert rows.location_code.duplicated().sum() == 1
    assert rows.location_code.isna().sum() == 1
    assert rows.loc[14, "zone"] == "UNKNOWN"
    assert rows.loc[15, "capacity_pallets"] == -1
    assert rows.loc[16, "capacity_pallets"] == "many"
    assert rows.loc[17, "location_type"] == "UNKNOWN"
    assert rows.loc[18, "location_code"] == "bad code"
    assert rows.loc[19, "capacity_pallets"] > RULES["max_capacity_pallets"]
    assert rows.iloc[:12].capacity_pallets.eq(10).all()


def test_database_has_reference_locations_and_customer_rules(tmp_path):
    _, database = generate(tmp_path)
    with sqlite3.connect(database) as connection:
        assert connection.execute("PRAGMA integrity_check").fetchone() == ("ok",)
        assert connection.execute("PRAGMA foreign_key_check").fetchall() == []
        assert connection.execute("SELECT customer_id FROM customers").fetchall() == [(CUSTOMER,)]
        codes = connection.execute("SELECT location_code FROM locations ORDER BY location_code").fetchall()
        assert codes == [("LOC-001",), ("LOC-100",), ("LOC-101",)]
        rules = {name: json.loads(value) for name, value in connection.execute(
            "SELECT rule_name, rule_value_json FROM customer_rules")}
        assert rules == RULES


@pytest.mark.parametrize("existing", ["warehouse_locations.xlsx", "synthetic_wms.sqlite3"])
def test_generator_preserves_existing_files(tmp_path, existing):
    path = tmp_path / existing
    path.write_bytes(b"existing local data")
    with pytest.raises(FileExistsError):
        generate(tmp_path)
    assert path.read_bytes() == b"existing local data"
    assert len(list(tmp_path.iterdir())) == 1
