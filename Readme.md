# Excel-to-csv-exporter

A fast .NET 8 command-line tool to export Excel `.xlsx` / `.xls` files
to CSV.\
Supports large files (50k+ rows × 200+ columns), custom sheet selection,
delimiters, and cell ranges.

------------------------------------------------------------------------

## 🚀 Features

-   Handles **large Excel files** efficiently (streaming read, no full
    in-memory load).\
-   Export from:
    -   **First sheet** (default)\
    -   **Specific sheet index** (zero-based)\
    -   **Specific sheet name**\
-   Choose your own **CSV delimiter** (comma, semicolon, pipe, tab,
    etc.).\
-   Export a specific **cell range** (e.g., `A1:D100`, `B2:Z50000`).\
-   Produces **UTF-8 CSV output**.

------------------------------------------------------------------------

## ⚙️ Usage

### Syntax

    Excel-to-csv-exporter <input.xlsx> <output.csv> [sheetName|sheetIndex] [delimiter] [range]

### Parameters

  ------------------------------------------------------------------------
  Parameter                      Required          Description
  ------------------------------ ----------------- -----------------------
  `<input.xlsx>`                 ✅ Yes            Path to the Excel file
                                                   to export

  `<output.csv>`                 ✅ Yes            Path to the CSV file to
                                                   create

  `[sheetName|sheetIndex]`       Optional          Sheet to export.
                                                   Default = first sheet.
                                                   `<br>`{=html}• Index is
                                                   zero-based
                                                   (`0 = first sheet`).
                                                   `<br>`{=html}• Or
                                                   specify sheet name
                                                   (e.g. `"Data2025"`)

  `[delimiter]`                  Optional          CSV field separator.
                                                   Default = `,`.
                                                   Examples: `";"`, `"|"`,
                                                   `"   "`

  `[range]`                      Optional          Cell range in Excel
                                                   style (inclusive).
                                                   Example: `A1:D100`.
  ------------------------------------------------------------------------

------------------------------------------------------------------------

## 📖 Examples

### Export first sheet with default comma

    Excel-to-csv-exporter big.xlsx output.csv

### Export second sheet (index = 1) with semicolon

    Excel-to-csv-exporter big.xlsx output.csv 1 ";"

### Export sheet by name with pipe, range A1:D100

    Excel-to-csv-exporter big.xlsx output.csv "January Data" "|" A1:D100

### Export rows 1000--2000 and cols B--Z from first sheet

    Excel-to-csv-exporter big.xlsx output.csv 0 "," B1000:Z2000

------------------------------------------------------------------------

## 🛠️ Build from Source

1.  Clone or copy this project.

2.  Install dependencies:

    ``` bash
    dotnet add package ExcelDataReader
    dotnet add package ExcelDataReader.DataSet
    ```

3.  Build:

    ``` bash
    dotnet build -c Release
    ```

    Output: `bin/Release/net8.0/Excel-to-csv-exporter.dll`

Run with:

``` bash
dotnet bin/Release/net8.0/Excel-to-csv-exporter.dll <args>
```

------------------------------------------------------------------------

## 📦 Publish as Standalone Binary

To create a self-contained `.exe` (no .NET runtime required):

``` bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

Output binary:

    bin/Release/net8.0/win-x64/publish/Excel-to-csv-exporter.exe

Run with:

``` bash
Excel-to-csv-exporter.exe big.xlsx out.csv "Sheet1" ";" A1:D100
```

For Linux/macOS, change `-r` to `linux-x64` or `osx-x64`.

------------------------------------------------------------------------

## ⚡ Notes

-   ExcelDataReader is used for **streaming reads** → keeps memory low.\
-   Quotes inside cells (`"`) are escaped as `""`.\
-   If you need full RFC 4180 quoting (wrap fields that contain
    delimiters/line breaks in quotes), you can extend the writer logic.
