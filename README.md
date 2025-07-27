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
        row ??= new object[excelFile.FieldCount];
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

### XLSB Read
| Method                              | FileName             | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated |
|------------------------------------ |--------------------- |----------:|---------:|---------:|----------:|----------:|----------:|----------:|
| 'SpreadSheetTasks - XLSB Read - v1' | 200kFile.xlsb        |  99.03 ms | 1.925 ms | 1.607 ms | 5400.0000 | 3800.0000 | 1400.0000 |  68.48 MB |
| 'SpreadSheetTasks - XLSB Read - v2' | 200kFile.xlsb        | 115.19 ms | 2.246 ms | 3.148 ms | 5000.0000 | 3500.0000 | 1000.0000 |  49.03 MB |
| 'Sylvan.Data.Excel - XLSB Read'     | 200kFile.xlsb        | 127.57 ms | 2.508 ms | 3.433 ms | 6000.0000 | 3000.0000 | 1000.0000 |  50.82 MB |
| 'SpreadSheetTasks - XLSB Read - v1' | 65K_R(...).xlsb [21] |  49.02 ms | 0.204 ms | 0.191 ms | 2363.6364 |  818.1818 |  727.2727 |  28.83 MB |
| 'SpreadSheetTasks - XLSB Read - v2' | 65K_R(...).xlsb [21] |  64.47 ms | 0.833 ms | 0.696 ms | 1666.6667 |         - |         - |  13.66 MB |
| 'Sylvan.Data.Excel - XLSB Read'     | 65K_R(...).xlsb [21] |  71.58 ms | 0.512 ms | 0.479 ms | 2875.0000 |  125.0000 |         - |  23.16 MB |


### XLSX Read
| Method               | Mean     | Error   | StdDev  | Gen0      | Gen1      | Gen2      | Allocated   |
|--------------------- |---------:|--------:|--------:|----------:|----------:|----------:|------------:|
| SpreadSheetTasks200K | 244.0 ms | 1.03 ms | 0.91 ms | 5000.0000 | 3000.0000 | 1000.0000 | 35040.13 KB |
| Sylvan200k           | 329.5 ms | 4.45 ms | 4.16 ms | 6000.0000 | 3000.0000 | 1000.0000 | 52319.76 KB |
| SpreadSheetTasks65k  | 170.4 ms | 1.46 ms | 1.37 ms |         - |         - |         - |   491.09 KB |
| Sylvan65K            | 172.8 ms | 1.13 ms | 1.06 ms |         - |         - |         - |   661.50 KB |


### XLSB Write
| Method                          | ReaderType | Mean     | Error   | StdDev  | Gen0      | Gen1      | Allocated |
|-------------------------------- |----------- |---------:|--------:|--------:|----------:|----------:|----------:|
| 'SpreadSheetTasks - XLSB Write' | GENERAL    | 116.6 ms | 1.60 ms | 1.50 ms | 1600.0000 |         - |  30.57 MB |
| XlsbSylvanWrite                 | GENERAL    | 162.4 ms | 2.27 ms | 2.12 ms | 1666.6667 |         - |  36.75 MB |


### XLSX Write
| Method                          | ReaderType | Mean     | Error   | StdDev  | Gen0      | Gen1      | Allocated |
|-------------------------------- |----------- |---------:|--------:|--------:|----------:|----------:|----------:|
| 'SpreadSheetTasks - XLSX Write' | GENERAL    | 171.1 ms | 2.24 ms | 1.99 ms | 1500.0000 |         - |  30.74 MB |
| 'Sylvan - XLSX Write'           | GENERAL    | 220.6 ms | 2.83 ms | 2.51 ms | 2500.0000 | 1000.0000 |  42.94 MB |


https://github.com/KrzysztofDusko/SpreadSheetTasks
 