::: {align="center"}
# SheetSync

### Move the right data. Keep your workbook intact.

**SheetSync** helps you copy selected columns from one spreadsheet into
another by matching rows through a shared key --- without manually
searching, copying, and pasting thousands of cells.

[Launch
SheetSync](https://sheetsyncgit-evgon2k5fuf2hhq7bl7p9g.streamlit.app) ·
[Report a bug](https://github.com/Aryamann687/SheetSync/issues) · [View
source](https://github.com/Aryamann687/SheetSync)
:::

------------------------------------------------------------------------

## ✨ What is SheetSync?

Have one spreadsheet with employee IDs and salaries, and another with
employee IDs and names? SheetSync matches the records and transfers the
columns you choose.

Instead of manually using VLOOKUP, XLOOKUP, or copy-paste, upload your
files, choose the matching columns, review the preview, and download the
updated destination workbook.

> **Destination-first workflow:** your destination workbook defines
> which records are updated. SheetSync does not add source-only records
> to the destination.

## 🚀 Features

-   **Simple, guided workflow** --- upload, select, preview, download.
-   **Column mapping** --- choose which source columns go into which
    destination columns.
-   **Flexible row matching** --- match records using shared key
    columns.
-   **Matching modes** --- use recommended matching or strict exact
    matching.
-   **Review before applying** --- inspect proposed changes and
    unmatched records.
-   **Safer updates** --- preserve existing values in fill-only-blank
    mode and protect formula cells.
-   **Original file stays untouched** --- changes are applied to an
    output copy.
-   **Useful reports** --- understand matched records, updates, and
    records that could not be matched.

## 🧭 How it works

1.  **Upload your source** --- the workbook or CSV containing the
    information you want to copy.
2.  **Upload your destination** --- the workbook you want to update.
3.  **Choose matching columns** --- identify the key column in each
    file.
4.  **Map columns** --- select the source columns and their destination
    counterparts.
5.  **Review the preview** --- inspect proposed changes and unmatched
    rows.
6.  **Download** --- save the updated destination file.

### Example

**Source file**

  Employee ID     Salary
  ------------- --------
  E101             50000
  E102             60000
  E103             45000

**Destination file**

  Employee ID   Name      Salary
  ------------- ------- --------
  E101          Aman    
  E102          Rahul      55000
  E103          Priya   

With *fill-only-blank* enabled, SheetSync fills E101 and E103 while
preserving Rahul's existing salary.

## 🛠️ Run locally

### Requirements

-   Python 3.10+ (developed and tested with Python 3.12)
-   Git

### Setup

``` bash
git clone https://github.com/Aryamann687/SheetSync.git
cd SheetSync

# Windows PowerShell
python -m venv venv
.\venv\Scripts\Activate.ps1

# macOS / Linux
# python3 -m venv venv
# source venv/bin/activate

python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Streamlit will open SheetSync in your browser. If it doesn't, follow the
local URL printed in your terminal.

## 🧪 Run tests

``` bash
python -m pytest -v
```

The project has a test suite covering matching, planning, consolidation,
and application flows.

## 🧱 Project structure

``` text
SheetSync/
├── app.py                 # Streamlit interface
├── core/
│   ├── reader.py          # Spreadsheet reading and preparation
│   ├── matcher.py         # Record matching logic
│   ├── updater.py         # Safe workbook update operations
│   └── consolidator.py    # Planning and consolidation engine
├── tests/                 # Automated tests
├── requirements.txt
└── README.md
```

## 🔐 Data and safety

-   SheetSync works on an output copy; it does not overwrite your
    uploaded destination file.
-   Formula cells are protected from unintended updates.
-   Matching and proposed changes should be reviewed before downloading.
-   Use dummy or non-sensitive files when trying the public demo.
-   Do not upload confidential, financial, employee, or personally
    identifiable data to a public deployment unless you have verified
    the hosting and data-handling requirements for your use case.

## 🗺️ Roadmap

-   [ ] Add missing destination columns during mapping
-   [ ] Expand the user interface for multi-source consolidation
-   [ ] Additional preview and reporting improvements

## 🤝 Contributing

Issues, ideas, and pull requests are welcome.

1.  Fork the repository.
2.  Create a feature branch.
3.  Make your changes and add tests where appropriate.
4.  Run `python -m pytest -v`.
5.  Open a pull request with a clear description.

## 📄 License

See the repository for license information. If no license file is
present, the project is not yet explicitly licensed for reuse.

------------------------------------------------------------------------

::: {align="center"}
**Built to make spreadsheet work less manual.**
:::
