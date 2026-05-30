# SpreadSheetTasks

The .NET library for fast reading and writing Excel files (.xlsx, .xlsb).


## Installation

```bash
dotnet add package SpreadSheetTasks
```

```xml
<PackageReference Include="SpreadSheetTasks" Version="0.6.1" />
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

## Benchmarks

Tested on: Windows 11, AMD Ryzen 7 7840HS, .NET 10.0.8 (10.0.8)

### XLSB Read
| Method                              | FileName             | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated |
|------------------------------------ |--------------------- |----------:|---------:|---------:|----------:|----------:|----------:|----------:|
| 'SpreadSheetTasks - XLSB Read - v1' | 200kFile.xlsb        | 106.01 ms | 2.088 ms | 3.922 ms | 5400.0000 | 3800.0000 | 1400.0000 |  68.59 MB |
| 'SpreadSheetTasks - XLSB Read - v2' | 200kFile.xlsb        | 113.97 ms | 2.264 ms | 5.380 ms | 5000.0000 | 3500.0000 | 1000.0000 |  49.13 MB |
| 'Sylvan.Data.Excel - XLSB Read'     | 200kFile.xlsb        | 129.04 ms | 2.820 ms | 8.137 ms | 5000.0000 | 2000.0000 | 1000.0000 |  50.82 MB |
| 'SpreadSheetTasks - XLSB Read - v1' | 65K_R(...).xlsb [21] |  51.08 ms | 0.608 ms | 0.539 ms | 2400.0000 |  800.0000 |  700.0000 |  28.93 MB |
| 'SpreadSheetTasks - XLSB Read - v2' | 65K_R(...).xlsb [21] |  56.47 ms | 1.055 ms | 0.987 ms | 1666.6667 |         - |         - |  13.76 MB |
| 'Sylvan.Data.Excel - XLSB Read'     | 65K_R(...).xlsb [21] |  63.22 ms | 0.469 ms | 0.415 ms | 2875.0000 |  125.0000 |         - |  23.16 MB |

### XLSX Read
| Method               | Mean     | Error   | StdDev   | Gen0      | Gen1      | Gen2      | Allocated   |
|--------------------- |---------:|--------:|---------:|----------:|----------:|----------:|------------:|
| SpreadSheetTasks200K | 267.5 ms | 6.57 ms | 19.18 ms | 5000.0000 | 3000.0000 | 1000.0000 | 35139.8 KB |
| Sylvan200k           | 334.4 ms | 5.79 ms |  5.42 ms | 6000.0000 | 3000.0000 | 1000.0000 | 52327.52 KB |
| SpreadSheetTasks65k  | 170.0 ms | 3.37 ms |  2.98 ms |         - |         - |         - |   593.92 KB |
| Sylvan65K            | 166.6 ms | 2.86 ms |  2.68 ms |         - |         - |         - |   664.77 KB |

### XLSB Write (200k rows, mixed types)
| Method                          | ReaderType | Mean     | Error   | StdDev  | Gen0      | Allocated |
|-------------------------------- |----------- |---------:|--------:|--------:|----------:|----------:|
| 'SpreadSheetTasks - XLSB Write' | GENERAL    | 127.1 ms | 3.33 ms | 9.76 ms | 1750.0000 |  30.57 MB |
| XlsbSylvanWrite                 | GENERAL    | 178.6 ms | 3.49 ms | 5.12 ms | 1000.0000 |  36.75 MB |

### XLSX Write (200k rows, mixed types)
| Method                          | ReaderType | Mean     | Error   | StdDev  | Gen0      | Allocated |
|-------------------------------- |----------- |---------:|--------:|--------:|----------:|----------:|
| 'SpreadSheetTasks - XLSX Write' | GENERAL    | 183.6 ms | 3.58 ms | 5.79 ms | 1500.0000 |  30.74 MB |

## Links

- NuGet: https://www.nuget.org/packages/SpreadSheetTasks/
- GitHub: https://github.com/KrzysztofDusko/SpreadSheetTasks
- License: MIT
