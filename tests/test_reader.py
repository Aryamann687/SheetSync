
from core.reader import list_sheets, load_table


def test_excel_reader():
    with open("sample.xlsx", "rb") as f:
        assert list_sheets(f) == ["Sheet1"]

    with open("sample.xlsx", "rb") as f:
        df = load_table(f, "Sheet1")

    assert "District" in df.columns
    assert "Block" in df.columns
    assert df["_xl_row"].tolist() == [2, 3]

import pandas as pd
from core.reader import list_sheets, load_table


def test_csv_reader():
    with open("sample.csv", "rb") as f:
        assert list_sheets(f) == ["CSV"]

    with open("sample.csv", "rb") as f:
        df = load_table(f, "CSV")

    assert df["District"].tolist() == ["Kolar", "Bengaluru"]
    assert df["_xl_row"].tolist() == [2, 3]


def test_custom_header_row():
    with open("header_test.xlsx", "rb") as f:
        df = load_table(f, "Sheet1", header_row=3)

    assert "District" in df.columns
    assert df["_xl_row"].tolist() == [4, 5]
