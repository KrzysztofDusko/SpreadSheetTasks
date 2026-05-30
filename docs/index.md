# SpreadSheetTasks Usage Guide

## Table of Contents

1. [Installation](#installation)
2. [Writing Excel Files](#writing-excel-files)
   - [Write from DataTable](#write-from-datatable)
   - [Write from List](#write-from-list)
   - [Write from string array](#write-from-string-array)
   - [Multi-sheet write](#multi-sheet-write)
   - [Write with Autofilter](#write-with-autofilter)
   - [Write with formatted cells](#write-with-formatted-cells)
   - [Write to Stream (MemoryStream / FileStream)](#write-to-stream)
   - [Hidden sheets](#hidden-sheets)
   - [Events: OnCompress / On10k](#events-oncompress--on10k)
   - [Document properties](#document-properties)
   - [SuppressSomeDate](#suppresssomedate)
3. [Reading Excel Files](#reading-excel-files)
   - [Basic read](#basic-read)
   - [Read with typed getters](#read-with-typed-getters)
   - [GetRowsOfSheet](#getrowsofsheet)
   - [Get sheet names](#get-sheet-names)
   - [RowCount and ResultsCount](#rowcount-and-resultscount)
   - [TreatAllColumnsAsText](#treatallcolumnsastext)
   - [UseMemoryStreamInXlsb](#usememorystreaminxlsb)
   - [Read in update mode (XLSX)](#read-in-update-mode-xlsx)
4. [Factory method](#factory-method-excelwritercreatewriter)
5. [Format constants (F class)](#format-constants-f-class)

---

## Installation

```bash
dotnet add package SpreadSheetTasks
```

```xml
<PackageReference Include="SpreadSheetTasks" Version="0.6.1" />
```

---

## Writing Excel Files

### Write from DataTable

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Name", typeof(string));
dt.Columns.Add("Age", typeof(int));
dt.Columns.Add("Salary", typeof(double));
dt.Rows.Add("Alice", 30, 5000.0);
dt.Rows.Add("Bob", 25, 4500.0);

using (var writer = new XlsxWriter("employees.xlsx"))
{
    writer.AddSheet("Employees");
    writer.WriteSheet(dt.CreateDataReader());
}

using (var writer = new XlsbWriter("employees.xlsb"))
{
    writer.AddSheet("Employees");
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Write from List

```csharp
using SpreadSheetTasks;

var headers = new List<string> { "Product", "Price", "Quantity" };
var types = new List<TypeCode> { TypeCode.String, TypeCode.Double, TypeCode.Int32 };
var rows = new List<object?[]>
{
    new object?[] { "Apple", 1.99, 100 },
    new object?[] { "Banana", 0.99, 250 },
    new object?[] { "Cherry", 3.49, 75 },
};

using (var writer = ExcelWriter.CreateWriter("products.xlsx"))
{
    writer.AddSheet("Products");
    writer.WriteSheet(headers, types, rows, headers: true, doAutofilter: true);
}
```

### Write from string array

```csharp
using SpreadSheetTasks;

string[] data = ["Alpha", "Beta", "Gamma"];

using (var writer = new XlsxWriter("strings.xlsx"))
{
    writer.AddSheet("Sheet1");
    writer.WriteSheet(data);
}
```

### Multi-sheet write

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Value", typeof(int));
dt.Rows.Add(1);
dt.Rows.Add(2);

using (var writer = ExcelWriter.CreateWriter("multi.xlsx"))
{
    writer.AddSheet("Sheet1");
    writer.WriteSheet(dt.CreateDataReader());

    writer.AddSheet("Sheet2");
    writer.WriteSheet(dt.CreateDataReader());

    writer.AddSheet("HiddenSheet", hidden: true);
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Write with Autofilter

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("City", typeof(string));
dt.Columns.Add("Population", typeof(int));
dt.Rows.Add("New York", 8_400_000);
dt.Rows.Add("Los Angeles", 3_800_000);
dt.Rows.Add("Chicago", 2_700_000);

using (var writer = ExcelWriter.CreateWriter("cities.xlsx"))
{
    writer.AddSheet("Cities");
    writer.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
}
```

### Write with formatted cells

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Description", typeof(string));
dt.Columns.Add("Value", typeof(object));

dt.Rows.Add("Thousands", new FormattedCell(1234567, F.THOUSANDS_SEP));
dt.Rows.Add("Currency PLN", new FormattedCell(1234.56, F.CURRENCY_PLN));
dt.Rows.Add("Date", new FormattedCell(new DateTime(2026, 6, 1), F.DATE_SHORT));
dt.Rows.Add("DateTime ISO", new FormattedCell(new DateTime(2026, 6, 1, 14, 30, 0), F.DATETIME_ISO));
dt.Rows.Add("Percentage", new FormattedCell(0.25, F.PERCENTAGE));
dt.Rows.Add("Scientific", new FormattedCell(12345.67, F.SCIENTIFIC));
dt.Rows.Add("Time", new FormattedCell(new DateTime(2026, 6, 1, 8, 15, 0), F.TIME_HH_MM_SS));

using (var writer = ExcelWriter.CreateWriter("formatted.xlsx"))
{
    writer.AddSheet("Formatted");
    writer.WriteSheet(dt.CreateDataReader());
}
```

You can also use custom format strings directly:

```csharp
dt.Rows.Add("Custom", new FormattedCell(1234.56, "#,##0.00 \"USD\""));
dt.Rows.Add("ZIP", new FormattedCell(12345, "00000"));
```

### Write to Stream

```csharp
using SpreadSheetTasks;
using System.Data;

// FileStream
using (var fs = File.Open("stream_output.xlsx", FileMode.Create))
using (var writer = new XlsxWriter(fs))
{
    writer.AddSheet("Sheet1");
    var dt = new DataTable();
    dt.Columns.Add("Col1", typeof(string));
    dt.Rows.Add("Hello");
    writer.WriteSheet(dt.CreateDataReader());
}

// MemoryStream
byte[] excelBytes;
using (var ms = new MemoryStream())
using (var writer = new XlsxWriter(ms))
{
    writer.AddSheet("Sheet1");
    var dt = new DataTable();
    dt.Columns.Add("Col1", typeof(string));
    dt.Rows.Add("World");
    writer.WriteSheet(dt.CreateDataReader());

    // leaveExcelArchiveOpen: true (default for stream)
    // data is available after Dispose
}
```

### Hidden sheets

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Data", typeof(string));
dt.Rows.Add("visible");
dt.Rows.Add("hidden data");

using (var writer = ExcelWriter.CreateWriter("hidden.xlsx"))
{
    writer.AddSheet("Visible");
    writer.WriteSheet(dt.CreateDataReader());

    writer.AddSheet("Hidden", hidden: true);
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Events: OnCompress / On10k

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Index", typeof(int));
for (int i = 0; i < 25_000; i++)
    dt.Rows.Add(i);

using (var writer = new XlsxWriter("events_demo.xlsx"))
{
    writer.OnCompress += () => Console.WriteLine("Starting compression...");
    writer.On10k += (row) => Console.WriteLine($"Written {row} rows...");

    writer.AddSheet("LargeData");
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Document properties

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Col1", typeof(string));
dt.Rows.Add("Data");

using (var writer = ExcelWriter.CreateWriter("docprop.xlsx"))
{
    writer.DocPopertyProgramName = "MyApplication";
    writer.AddSheet("Sheet1");
    writer.WriteSheet(dt.CreateDataReader());
}
```

### SuppressSomeDate

Suppresses DateTime values where the year is 1000:

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Date", typeof(DateTime));
dt.Rows.Add(new DateTime(1000, 1, 1)); // will be suppressed
dt.Rows.Add(new DateTime(2026, 6, 1)); // normal

using (var writer = ExcelWriter.CreateWriter("suppress.xlsx"))
{
    writer.SuppressSomeDate = true;
    writer.AddSheet("Sheet1");
    writer.WriteSheet(dt.CreateDataReader());
}
```

---

## Reading Excel Files

### Basic read

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("file.xlsx"); // auto-detects xlsx/xlsb by extension
    reader.ActualSheetName = "Sheet1";

    object[]? row = null;
    while (reader.Read())
    {
        row ??= new object[reader.FieldCount];
        reader.GetValues(row);
        // process row...
    }
}
```

### Read with typed getters

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    reader.ActualSheetName = "Sheet1";

    while (reader.Read())
    {
        string name = reader.GetString(0);
        int age = reader.GetInt32(1);
        double salary = reader.GetDouble(2);
        DateTime? hireDate = null;

        // GetDateTime throws InvalidCastException if cell is not a DateTime
        try { hireDate = reader.GetDateTime(3); }
        catch (InvalidCastException) { }
    }
}
```

### GetRowsOfSheet

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    int totalRows = reader.GetRowsOfSheet("Sheet1").Count();
}
```

### Get sheet names

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("multi.xlsx");
    string[] sheetNames = reader.GetScheetNames();

    foreach (string name in sheetNames)
    {
        reader.ActualSheetName = name;
        Console.WriteLine($"Sheet: {name}");
    }
}
```

### RowCount and ResultsCount

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    reader.ActualSheetName = "Sheet1";

    int rowCount = reader.RowCount;        // estimated row count
    int sheetCount = reader.ResultsCount;  // number of sheets
}
```

### TreatAllColumnsAsText

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    reader.TreatAllColumnsAsText = true; // all values returned as strings
    reader.ActualSheetName = "Sheet1";

    while (reader.Read())
    {
        string val = reader.GetString(0);
    }
}
```

### UseMemoryStreamInXlsb

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.UseMemoryStreamInXlsb = false; // lower RAM usage, slightly slower
    reader.Open("data.xlsb");
    reader.ActualSheetName = "Sheet1";

    while (reader.Read()) { /* ... */ }
}
```

### Read in update mode (XLSX only)

```csharp
using SpreadSheetTasks;
using System.Data;

// Open existing file in update mode and replace sheet data
using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("existing.xlsx", updateMode: true);

    var newData = new DataTable();
    newData.Columns.Add("Name", typeof(string));
    newData.Columns.Add("Score", typeof(int));
    newData.Rows.Add("Alice", 95);
    newData.Rows.Add("Bob", 87);

    string range = reader.ReplaceSheetData("Sheet1", newData.CreateDataReader());
    Console.WriteLine($"Replaced range: {range}");
}
```

---

## Factory method (ExcelWriter.CreateWriter)

```csharp
using SpreadSheetTasks;
using System.Data;

// Automatically selects XlsxWriter or XlsbWriter based on file extension
using (var writer = ExcelWriter.CreateWriter("data.xlsx"))
{
    writer.AddSheet("Sheet1");
    var dt = new DataTable();
    dt.Columns.Add("Col1", typeof(string));
    dt.Rows.Add("Hello");
    writer.WriteSheet(dt.CreateDataReader());
}
```

---

## Format constants (F class)

| Constant | Format String | Description |
|----------|---------------|-------------|
| `F.THOUSANDS_SEP` | `#,##0` | Number with thousands separator |
| `F.CURRENCY_PLN` | `#,##0.00 "zł"` | Polish currency |
| `F.CURRENCY_EUR` | `#,##0.00 €` | Euro currency |
| `F.PERCENTAGE` | `0%` | Percentage |
| `F.SCIENTIFIC` | `0.00E+00` | Scientific notation |
| `F.TWO_DECIMALS` | `#,##0.00` | Number with 2 decimals |
| `F.TEXT` | `@` | Text format |
| `F.LEADING_ZEROS` | `000000000` | Leading zeros (9 digits) |
| `F.DATE_SHORT` | `dd.mm.yyyy` | Short date |
| `F.DATE_LONG` | `d mmmm yyyy` | Long date |
| `F.DATE_ISO` | `yyyy-mm-dd` | ISO date |
| `F.DATETIME_SHORT` | `dd.mm.yyyy hh:mm` | Short date/time |
| `F.DATETIME_LONG` | `d mmmm yyyy hh:mm:ss` | Long date/time |
| `F.DATETIME_ISO` | `yyyy-mm-dd"T"hh:mm:ss` | ISO date/time |
| `F.TIME_HH_MM` | `hh:mm` | Time (hours:minutes) |
| `F.TIME_HH_MM_SS` | `hh:mm:ss` | Time (hours:minutes:seconds) |
| `F.TIME_12H` | `h:mm AM/PM` | 12-hour time |
| `F.TIME_MS` | `hh:mm:ss.000` | Time with milliseconds |
