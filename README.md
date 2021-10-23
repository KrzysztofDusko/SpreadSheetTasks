# SpreadSheetTasks
 The .NET library for fast reading and writing Excel files (.xlsx, .xlsb). 
 Some methods/ideas based on great libraries : 
 * https://github.com/MarkPflug/Sylvan.Data.Excel
 * https://github.com/MarkPflug/Sylvan
 * https://github.com/ExcelDataReader/ExcelDataReader

 ## Installation
 https://www.nuget.org/packages/SpreadSheetTasks/
 
 ```Install-Package SpreadSheetTasks```
 
 ```dotnet add package SpreadSheetTasks```

 
 ## Usage
 
 ### Read
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
 ### Write
 ``` C#
using (XlsbWriter xlsx = new XlsbWriter("file.xlsb"))
{
   xlsx.AddSheet("sheetName");
   xlsx.WriteSheet(dataReader);
}
 ```
 
 ## Benchamarks

 https://github.com/KrzysztofDusko/SpreadSheetTasks/tree/main/source/Benchmark
 
 ``` ini
BenchmarkDotNet=v0.13.1, OS=Windows 10.0.22000
Intel Core i5-7500 CPU 3.40GHz (Kaby Lake), 1 CPU, 4 logical and 4 physical cores
.NET SDK=6.0.100-rc.2.21505.57
  [Host]   : .NET 6.0.0 (6.0.21.48005), X64 RyuJIT
  .NET 5.0 : .NET 5.0.11 (5.0.1121.47308), X64 RyuJIT
  .NET 6.0 : .NET 6.0.0 (6.0.21.48005), X64 RyuJIT
```
### Read
Xlsx

|               Method |  Runtime |     Mean |   Error |  StdDev |     Gen 0 |     Gen 1 |     Gen 2 | Allocated |
|--------------------- |--------- |---------:|--------:|--------:|----------:|----------:|----------:|----------:|
| SpreadSheetTasks200K | .NET 5.0 | 731.0 ms | 2.88 ms | 2.25 ms | 6000.0000 | 2000.0000 | 1000.0000 |     34 MB |
| SpreadSheetTasks200K | .NET 6.0 | 662.9 ms | 5.80 ms | 5.14 ms | 6000.0000 | 2000.0000 | 1000.0000 |     34 MB |


Xlsb
|   Method |  Runtime |      FileName | InMemory |     Mean |   Error |  StdDev |     Gen 0 |     Gen 1 |     Gen 2 | Allocated |
|--------- |--------- |-------------- |--------- |---------:|--------:|--------:|----------:|----------:|----------:|----------:|
| ReadFile | .NET 5.0 | 200kFile.xlsb |    False | 290.8 ms | 1.98 ms | 1.75 ms | 6000.0000 | 2000.0000 | 1000.0000 |     34 MB |
| ReadFile | .NET 6.0 | 200kFile.xlsb |    False | 249.7 ms | 1.05 ms | 0.98 ms | 6000.0000 | 2000.0000 | 1000.0000 |     34 MB |
| ReadFile | .NET 5.0 | 200kFile.xlsb |     True | 214.7 ms | 2.57 ms | 2.41 ms | 8000.0000 | 4333.3333 | 1333.3333 |     68 MB |
| ReadFile | .NET 6.0 | 200kFile.xlsb |     True | 195.6 ms | 0.54 ms | 0.48 ms | 8000.0000 | 4333.3333 | 1333.3333 |     68 MB |


### Write
|             Method |  Runtime |   Rows |       Mean |    Error |  StdDev |     Gen 0 | Allocated |
|------------------- |--------- |------- |-----------:|---------:|--------:|----------:|----------:|
|   XlsxWriteDefault | .NET 5.0 | 200000 | 1,181.4 ms |  2.42 ms | 2.14 ms | 4000.0000 |     31 MB |
| XlsxWriteLowMemory | .NET 5.0 | 200000 | 1,157.4 ms | 10.21 ms | 9.55 ms | 4000.0000 |     14 MB |
|   XlsbWriteDefault | .NET 5.0 | 200000 |   687.9 ms |  2.41 ms | 2.25 ms | 4000.0000 |     31 MB |
|   XlsxWriteDefault | .NET 6.0 | 200000 | 1,160.1 ms |  1.42 ms | 1.26 ms | 4000.0000 |     31 MB |
| XlsxWriteLowMemory | .NET 6.0 | 200000 | 1,142.3 ms | 10.48 ms | 9.80 ms | 4000.0000 |     14 MB |
|   XlsbWriteDefault | .NET 6.0 | 200000 |   680.5 ms |  1.47 ms | 1.37 ms | 4000.0000 |     31 MB |