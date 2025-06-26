# TableToSheetDataBridge

📋 Copy Table Data via Mapping (Excel VBA Tool)

This Excel VBA macro automates the task of copying data from structured Excel Tables (ListObjects) into another sheet, using configurable mapping rules. It's perfect for extracting specific columns, transforming date formats, and sending data to a different layout — all with just one macro run.
---

✅ Features

✔️ Copy from Excel Tables by name and column (by index or header)

✔️ Paste to any sheet/column, starting from any row

✔️ Optional formatting (e.g., convert dates to `yyyymmdd`)

✔️ Central mapping configuration for full control

✔️ Reusable & readable structure
---

🔧 How It Works

The macro relies on a configuration function GetMappings() where each mapping defines:
```
Array(
  "SourceSheet",     ' Name of the sheet containing the source table
  "TableName",       ' Excel Table (ListObject) name
  ColumnID,           ' Header name (string) OR column number (1-based index)
  "TargetSheet",     ' Sheet to paste into
  "TargetColumn",    ' Column letter (e.g., "A", "B")
  StartRow,           ' Row to start pasting into (e.g., 3)
  Optional Format     ' e.g. "yyyymmdd" for dates (optional)
)
```
---

✨ Example
```
GetMappings = Array( _
    Array("MainData", "MainTable", 1, "DataToGo", "A", 3, "yyyymmdd"), _
    Array("MainData", "MainTable", 2, "DataToGo", "B", 3) _
)
```
This would:

Copy column 1 from the MainTable in MainData, format it as yyyymmdd, and paste into column A of DataToGo, starting at row 3.

Copy column 2 from the same table and paste it as-is into column B.
---

💾 Installation

Press `ALT + F11` in Excel to open the VBA Editor.

Insert a new module (Right-click project > Insert > Module).

Paste in the macro script.

Save your file as .xlsm (macro-enabled workbook).
---

▶️ How to Use

Adjust the `GetMappings()` function to define what you want to extract.

Run the macro `CopyMappingsData()`.

All mapped data will be copied and formatted into the destination sheet.

You can also assign the macro to a button or ribbon shortcut for quick access.
---

📌 Use Cases

Extracting specific table columns into upload-ready formats

Formatting dates and values without manual copy/paste

Standardizing exports from different input sheets
---

📋 Example Project Folder Structure
```
YourWorkbook.xlsm
├── Sheet1 (MainData)
│   └── Table: MainTable
├── Sheet2 (DataToGo)
└── VBA Module with this tool
```
---
🧠 Notes

The macro works with Excel Tables (ListObjects) only — not regular ranges.

You can use column names or indexes (e.g., "Date" or 1).

Use Format(...) strings like `"dd-mm-yyyy", "yyyymmdd"` to control output.

Target columns must be expressed in letters, not numbers (e.g., "A").
---

❗ Prerequisites

Excel 2016 or newer (with VBA support)

Macros must be enabled

Source data must be in Tables, not free-form ranges
---

📄 License

MIT License — use freely, contribute back if helpful 💙
---

👏 Author

Created by Mohamed El-ansary. This tool was built to help with structured data transformations in Excel workflows.
