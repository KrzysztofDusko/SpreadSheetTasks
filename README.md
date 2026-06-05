# SpreadSheetTasks

The .NET library for fast reading and writing Excel files (.xlsx, .xlsb).


## Installation

```bash
dotnet add package SpreadSheetTasks
```

```xml
<PackageReference Include="SpreadSheetTasks" Version="1.0.0" />
```

## Quick Start

### Write to Excel

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Name", typeof(string));
dt.Columns.Add("Age", typeof(int));
dt.Rows.Add("Alice", 30);
dt.Rows.Add("Bob", 25);

using (var writer = new XlsxWriter("output.xlsx")) // or XlsbWriter
{
    writer.AddSheet("People");
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Read from Excel

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("output.xlsx");
    reader.ActualSheetName = "People";

    object[]? row = null;
    while (reader.Read())
    {
        row ??= new object[reader.FieldCount];
        reader.GetValues(row);
    }
}
```

## Documentation

Full usage guide with detailed examples:
- [docs/index.md](docs/index.md)

### Write from object[][] (typed headers)

```csharp
using SpreadSheetTasks;

using (var writer = new XlsxWriter("products.xlsx"))
{
    writer.AddSheet("Products");
    writer.WriteSheet(
        rows: new object[][]
        {
            new object[] { "Apple", 1.99, 100 },
            new object[] { "Banana", 0.99, 250 },
        },
        headers: new[] { "Product", "Price", "Quantity" },
        headers_row: true,
        doAutofilter: true
    );
}
```

### Write with per-cell formatting (FormattedCell[][])

```csharp
using SpreadSheetTasks;

using (var writer = new XlsxWriter("formatted.xlsx"))
{
    writer.AddSheet("Sheet1");
    writer.WriteSheet(
        rows: new object[][]
        {
            new object[] { 1234567, 1234.56 },
            new object[] { 0.25, 12345.67 },
        },
        headers: new[] { "Thousands", "Scientific" },
        formats: new FormattedCell[][]
        {
            new FormattedCell[] { new(1234567, F.THOUSANDS_SEP), new(1234.56, F.CURRENCY_PLN) },
            new FormattedCell[] { new(0.25, F.PERCENTAGE), new(12345.67, F.SCIENTIFIC) },
        }
    );
}
```

### GetExcelDataType / GetNativeValue

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    reader.ActualSheetName = "Sheet1";

    while (reader.Read())
    {
        // Use typed getters for general use
        string name = reader.GetString(0);
        long value = reader.GetInt64(1);

        // Use GetExcelDataType to inspect cell type
        ExcelDataType type = reader.GetExcelDataType(2);

        // High-performance escape hatch (check type first!)
        if (type == ExcelDataType.Double)
        {
            double raw = reader.GetNativeValue(2).doubleValue;
        }
    }
}
```

### Update existing XLSX (ReplaceSheetData + ReplacePivotTableDim)

```csharp
using SpreadSheetTasks;
using System.Data;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("existing.xlsx", updateMode: true);

    var newData = new DataTable();
    newData.Columns.Add("Product", typeof(string));
    newData.Columns.Add("Revenue", typeof(decimal));
    newData.Rows.Add("Widget", 50000m);

    // Replace data in a sheet
    string range = reader.ReplaceSheetData("Sheet1", newData.CreateDataReader());

    // Update pivot table data source reference
    reader.ReplacePivotTableDim("PivotTable1", range);
}
```

### Write to existing Stream (XlsxWriter.WriteToExisting)

```csharp
using SpreadSheetTasks;
using System.Data;
using System.IO;
using System.IO.Compression;

// Open an existing xlsx as a ZipArchive and write into a sheet stream
using (var archive = ZipFile.Open("existing.xlsx", ZipArchiveMode.Update))
{
    var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
    using var writer = new StreamWriter(entry.Open());

    var dt = new DataTable();
    dt.Columns.Add("Name", typeof(string));
    dt.Rows.Add("Alice");

    int rowsWritten = XlsxWriter.WriteToExisting(writer, dt.CreateDataReader());
}
```

### Stream writer with custom buffer size

```csharp
using SpreadSheetTasks;
using System.Data;

// Larger buffer for better write throughput on large files
using (var writer = new XlsxWriter("large.xlsx", bufferSize: 65536))
{
    writer.AddSheet("Data");
    var dt = new DataTable();
    dt.Columns.Add("Col", typeof(string));
    for (int i = 0; i < 100_000; i++) dt.Rows.Add($"Row{i}");
    writer.WriteSheet(dt.CreateDataReader());
}
```

## Breaking Changes in v1.0.0

| Old (v0.6.1) | New (v1.0.0) |
|---|---|
| `GetScheetNames()` (typo) | `GetSheetNames()` |
| `DocPopertyProgramName` (typo) | `DocPropertyProgramName` |
| `SuppressSomeDate` | `SuppressYear1000Dates` |
| `overLimit` parameter | `maxRows` parameter |
| `RowCount` returns `123123123` for unknown | `RowCount` returns `-1` for unknown |
| `FormattingStreamWriter` (public) | `FormattingStreamWriter` (internal) |
| `UseMemoryStreamInXlsb` (field) | `UseMemoryStreamInXlsb` (property) |
| Duplicate `F.SHORT_DATE` etc. constants | Marked `[Obsolete]` |

Old names still compile with `[Obsolete]` warnings.

## Benchmarks

### Windows 11 (25H2), AMD Ryzen 7 7840HS, .NET 10.0.8, BenchmarkDotNet 0.15.8

#### XLSB Read (65k rows)
| Method                                      | Mean     | Error     | StdDev  | Gen0      | Gen1     | Gen2     | Allocated |
|-------------------------------------------- |---------:|----------:|--------:|----------:|---------:|---------:|----------:|
| 'SpreadSheetTasks - XLSB Read - v1' (quick)   | 52.56 ms |  3.121 ms | 0.17 ms | 2400.0000 | 800.0000 | 700.0000 | 28.93 MB |
| 'SpreadSheetTasks - XLSB Read - v2' (quick)   | 60.61 ms | 67.866 ms | 3.72 ms | 1666.6667 |        - |        - | 13.76 MB |

#### XLSX Read (65k rows, typed getters)
| Method              | Mean     | Error    | StdDev  | Allocated |
|-------------------- |---------:|---------:|--------:|----------:|
| SpreadSheetTasks65k | 178.2 ms | 124.7 ms | 6.83 ms | 593.92 KB |

#### XLSB Write (50k rows, mixed types)
| Method                          | ReaderType | Mean     | Error     | StdDev  | Gen0      | Gen1     | Gen2     | Allocated |
|-------------------------------- |----------- |---------:|----------:|--------:|----------:|---------:|---------:|----------:|
| 'SpreadSheetTasks - XLSB Write' | GENERAL    | 40.40 ms | 13.952 ms | 0.77 ms |  916.6667 | 166.6667 | 83.3333 | 10.86 MB |
| XlsbSylvanWrite                 | GENERAL    | 51.34 ms | 12.901 ms | 0.71 ms |  545.4545 | 181.8182 | 90.9091 |  8.98 MB |

#### XLSX Write (50k rows, mixed types)
| Method                          | ReaderType | Mean     | Error     | StdDev  | Gen0      | Gen1      | Gen2     | Allocated |
|-------------------------------- |----------- |---------:|----------:|--------:|----------:|----------:|--------:|----------:|
| 'SpreadSheetTasks - XLSX Write' | GENERAL    | 58.08 ms |  8.808 ms | 0.48 ms | 1111.1111 |  111.1111 |       - | 13.31 MB |
| XlsxSylvanWrite                 | GENERAL    | 73.63 ms | 73.330 ms | 4.02 ms |  571.4286 |  142.8571 |       - | 10.51 MB |

### macOS Tahoe 26.5.1 (Apple M4), .NET 10.0.8, BenchmarkDotNet 0.15.8

#### XLSB Read (65k rows)
| Method                                      | Mean     | Error     | StdDev  | Gen0      | Gen1     | Gen2     | Allocated |
|-------------------------------------------- |---------:|----------:|--------:|----------:|---------:|---------:|----------:|
| 'SpreadSheetTasks - XLSB Read - v1' (quick)   | 40.56 ms |  2.895 ms | 1.59 ms | 2461.5385 | 846.1538 | 769.2308 |  28.93 MB |
| 'SpreadSheetTasks - XLSB Read - v2' (quick)   | 47.62 ms |  3.617 ms | 1.98 ms | 1666.6667 |        - |        - |  13.76 MB |

#### XLSX Read (65k rows, typed getters)
| Method              | Mean     | Error    | StdDev  | Allocated |
|-------------------- |---------:|---------:|--------:|----------:|
| SpreadSheetTasks65k | 152.3 ms |  1.29 ms | 1.01 ms |  593.88 KB |

#### XLSB Write (50k rows, mixed types)
| Method                          | ReaderType | Mean     | Error     | StdDev  | Gen0      | Gen1     | Gen2     | Allocated |
|-------------------------------- |----------- |---------:|----------:|--------:|----------:|---------:|---------:|----------:|
| 'SpreadSheetTasks - XLSB Write'  | GENERAL    | 23.97 ms |  0.178 ms | 0.166 ms |  968.7500 | 187.5000 | 93.7500 |  10.86 MB |
| XlsbSylvanWrite                 | GENERAL    | 32.48 ms |  0.509 ms | 0.425 ms |  562.5000 | 187.5000 | 125.0000 |   8.98 MB |

#### XLSX Write (50k rows, mixed types)
| Method                          | ReaderType | Mean     | Error     | StdDev  | Gen0      | Gen1      | Gen2     | Allocated |
|-------------------------------- |----------- |---------:|----------:|--------:|----------:|----------:|--------:|----------:|
| 'SpreadSheetTasks - XLSX Write' | GENERAL    | 41.43 ms |  0.781 ms | 0.731 ms | 1230.7692 | 230.7692 |        - |  13.31 MB |
| XlsxSylvanWrite                 | GENERAL    | 41.18 ms |  0.495 ms | 0.386 ms |  750.0000 | 250.0000 | 83.3333 |  10.51 MB |

## Links

- NuGet: https://www.nuget.org/packages/SpreadSheetTasks/
- GitHub: https://github.com/KrzysztofDusko/SpreadSheetTasks
- License: MIT
