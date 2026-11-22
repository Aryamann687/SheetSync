
import pandas as pd


def list_sheets(file):
    file.seek(0)

    if file.name.lower().endswith(".csv"):
        return ["CSV"]

    excel = pd.ExcelFile(file)
    return excel.sheet_names


def load_table(file, sheet=None, header_row=1):
    if header_row < 1:
        raise ValueError("Header row must be at least 1")

    file.seek(0)
    filename = file.name.lower()

    if filename.endswith(".csv"):
        try:
            df = pd.read_csv(
                file, header=header_row - 1,
                encoding="utf-8-sig"
            )
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(
                file, header=header_row - 1,
                encoding="cp1252"
            )

    elif filename.endswith((".xlsx", ".xls")):
        df = pd.read_excel(
            file,
            sheet_name=sheet if sheet is not None else 0,
            header=header_row - 1
        )

    else:
        raise ValueError("Unsupported file format")

    df.columns = [str(col).strip() for col in df.columns]
    if df.empty:
     raise ValueError("The selected file or sheet contains no data.")

    if any(not col or col.startswith("Unnamed:") for col in df.columns):
     raise ValueError("One or more column headers are missing.")

    if df.columns.duplicated().any():
     raise ValueError("Duplicate column headers found.")

    if "_xl_row" in df.columns:
     raise ValueError("Reserved column name: _xl_row")

    if "_xl_row" in df.columns:
        raise ValueError("Reserved column name: _xl_row")

    df["_xl_row"] = range(
        header_row + 1,
        header_row + len(df) + 1
    )

    return df
