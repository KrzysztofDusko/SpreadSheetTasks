# SpreadSheetTasks
 The .NET library for fast reading and writing Excel files (.xlsx, .xlsb)
 
 ## Installation
 https://www.nuget.org/packages/SpreadSheetTasks/
 
 ```Install-Package SpreadSheetTasks -Version 0.0.1```
 
 ```dotnet add package SpreadSheetTasks --version 0.0.1```

 

 ## Usage
 
 ```c#
 using (XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit())
 {
    excelFile.Open("file.xlsx");
    excelFile.ActualSheetName = "sheet1";
    object[] row = null;
    while (excelFile.Read())
    {
        if (row == null)
        {
            row = new object[excelFile.FieldCount];
        }
        excelFile.GetValues(row);
    }
 }
 ```
 
 ## Benchamarks
 ``` ini

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.19043.1237 (21H1/May2021Update)
Intel Core i5-7500 CPU 3.40GHz (Kaby Lake), 1 CPU, 4 logical and 4 physical cores
.NET SDK=5.0.401
  [Host]     : .NET 5.0.10 (5.0.1021.41214), X64 RyuJIT
  DefaultJob : .NET 5.0.10 (5.0.1021.41214), X64 RyuJIT
```
### Read
## Xlsx

|   Method |      FileName |     Mean |   Error |  StdDev |      Gen 0 |     Gen 1 |     Gen 2 | Allocated |
|--------- |-------------- |---------:|--------:|--------:|-----------:|----------:|----------:|----------:|
| **ReadFile** | **200kFile.xlsx** | **797.0 ms** | **5.08 ms** | **4.24 ms** |  **8000.0000** | **4000.0000** | **2000.0000** |     **38 MB** |

## Xlsb
|   Method |      FileName | UseMemoryStreamInXlsb |     Mean |   Error |  StdDev |     Gen 0 |     Gen 1 |     Gen 2 | Allocated |
|--------- |-------------- |---------------------- |---------:|--------:|--------:|----------:|----------:|----------:|----------:|
| ReadFile | 200kFile.xlsb |                 False | 281.7 ms | 3.10 ms | 2.90 ms | 8000.0000 | 4000.0000 | 2000.0000 |     38 MB |
| ReadFile | 200kFile.xlsb |                  True | 207.9 ms | 0.69 ms | 0.61 ms | 8000.0000 | 4666.6667 | 1666.6667 |     72 MB |

### Write
to do
