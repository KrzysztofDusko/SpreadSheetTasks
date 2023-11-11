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
 
 ## Benchamarks and more

### XLSB Read
| Method                              | FileName             | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated |
|------------------------------------ |--------------------- |----------:|---------:|---------:|----------:|----------:|----------:|----------:|
| 'SpreadSheetTasks - XLSB Read - v1' | 200kFile.xlsb        | 117.29 ms | 1.907 ms | 1.784 ms | 3400.0000 | 2800.0000 | 1400.0000 |  68.49 MB |
| 'SpreadSheetTasks - XLSB Read - v2' | 200kFile.xlsb        | 138.62 ms | 2.752 ms | 2.826 ms | 3000.0000 | 2000.0000 | 1000.0000 |  49.03 MB |
| 'Sylvan.Data.Excel - XLSB Read'     | 200kFile.xlsb        | 147.70 ms | 1.693 ms | 1.500 ms | 3000.0000 | 2500.0000 | 1500.0000 |  50.82 MB |
| 'SpreadSheetTasks - XLSB Read - v1' | 65K_R(...).xlsb [21] |  60.25 ms | 0.504 ms | 0.447 ms | 1555.5556 |  777.7778 |  777.7778 |  28.83 MB |
| 'SpreadSheetTasks - XLSB Read - v2' | 65K_R(...).xlsb [21] |  75.96 ms | 0.346 ms | 0.323 ms |  666.6667 |         - |         - |  13.66 MB |
| 'Sylvan.Data.Excel - XLSB Read'     | 65K_R(...).xlsb [21] |  90.80 ms | 0.641 ms | 0.535 ms | 1000.0000 |         - |         - |  23.16 MB |

### XLSB Write
| Method                          | ReaderType | Mean     | Error   | StdDev  | Gen0     | Allocated |
|-------------------------------- |----------- |---------:|--------:|--------:|---------:|----------:|
| 'SpreadSheetTasks - XLSB Write' | GENERAL    | 178.8 ms | 1.25 ms | 1.11 ms | 500.0000 |  30.57 MB |
| XlsbSylvanWrite                 | GENERAL    | 233.3 ms | 1.14 ms | 1.01 ms |        - |  36.75 MB |

https://github.com/KrzysztofDusko/SpreadSheetTasks
 