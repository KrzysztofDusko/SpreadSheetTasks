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
    object[]? row = null;
    while (excelFile.Read())
    {
        row ??= new object[excelFile.FieldCount];
        excelFile.GetValues(row);
    }
 }
 ```
 ### Write from IDataReader (best performance)
 ``` C#
using (XlsbWriter xlsx = new XlsbWriter("file.xlsb"))
{
   xlsx.AddSheet("sheetName");
   xlsx.WriteSheet(dataReader);
}
 ```
 

 ### Write from List
 ``` C#
// Write data from Lists
List<string> headers = new List<string> { "col1", "col2" };
List<TypeCode> typeCodes = new List<TypeCode> { TypeCode.Int32, TypeCode.String };
List<object[]> data = [ [1,"x"], [2,"y"], [3,"z"], [4,null], [null,"dda"]];

using (XlsbWriter writer = new XlsbWriter("output.xlsb"))
{
    writer.AddSheet("sheetName");
    writer.WriteSheet(headers, typeCodes, data, doAutofilter: true);
}
 ```

### Write from DataTable
```C#
// Create sample DataTable
DataTable dataTable = new DataTable();
dataTable.Columns.Add("COL1_INT", typeof(int));
dataTable.Columns.Add("COL2_TXT", typeof(string));
dataTable.Columns.Add("COL3_DATETIME", typeof(DateTime));
dataTable.Columns.Add("COL4_DOUBLE", typeof(double));

// Add some data
dataTable.Rows.Add(1, "Text1", DateTime.Now, 1.23);
dataTable.Rows.Add(2, "Text2", DateTime.Now.AddDays(1), 4.56);

// Write to Excel
using (var excel = new XlsxWriter("output.xlsx"))
{
    excel.AddSheet("sheetName");
    excel.WriteSheet(dataTable.CreateDataReader(), doAutofilter: true);
}
```

### More Examples

#### Write Multiple Sheets
```C#
using (var excel = new XlsxWriter("multisheet.xlsx"))
{
    excel.AddSheet("Sheet1");
    excel.WriteSheet(dataReader1, doAutofilter: true);
    
    excel.AddSheet("Sheet2");
    excel.WriteSheet(dataReader2, doAutofilter: true);
}
```

#### Write to MemoryStream
```C#
using (var memoryStream = new MemoryStream())
{
    var excel = new XlsxWriter(memoryStream);
    excel.AddSheet("Sheet1");
    excel.WriteSheet(dataReader, doAutofilter: true);
    excel.Dispose();

    // Use the MemoryStream
    memoryStream.Seek(0, SeekOrigin.Begin);
    // Save to file if needed
    using (var fileStream = File.Open("output.xlsb", FileMode.Create))
    {
        memoryStream.CopyTo(fileStream);
    }
}
```

#### XLSX vs XLSB Format
```C#
// format based on the file extension 
using (var excel = ExcelWriter.CreateWriter("file_path"))
{
    excel.AddSheet("Sheet1");
    excel.WriteSheet(dataReader);
}

// XLSX format
using (var excel = new XlsxWriter("file.xlsx"))
{
    excel.AddSheet("Sheet1");
    excel.WriteSheet(dataReader);
}

// XLSB format (better performance)
using (var excel = new XlsbWriter("file.xlsb"))
{
    excel.AddSheet("Sheet1");
    excel.WriteSheet(dataReader);
}

```

#### Read Sheet Names
```C#
using (var excelFile = new XlsxOrXlsbReadOrEdit())
{
    excelFile.Open("file.xlsx");
    var sheetNames = excelFile.GetScheetNames();
    foreach (var sheetName in sheetNames)
    {
        Console.WriteLine(sheetName);
    }
}
```

## Benchamarks and more


### XLSX Read
| Method               | Mean     | Error   | StdDev  | Gen0      | Gen1      | Gen2      | Allocated   |
|--------------------- |---------:|--------:|--------:|----------:|----------:|----------:|------------:|
| SpreadSheetTasks200K | 257.9 ms | 1.13 ms | 1.06 ms | 5000.0000 | 3000.0000 | 1000.0000 | 35038.91 KB |
| Sylvan200k           | 328.8 ms | 5.71 ms | 5.07 ms | 6000.0000 | 3000.0000 | 1000.0000 | 52321.90 KB |
| SpreadSheetTasks65k  | 181.7 ms | 1.87 ms | 1.66 ms |         - |         - |         - |   493.05 KB |
| Sylvan65K            | 175.8 ms | 3.44 ms | 5.56 ms |         - |         - |         - |   658.82 KB |


### XLSB Read
| Method                              | FileName             | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated |
|------------------------------------ |--------------------- |----------:|---------:|---------:|----------:|----------:|----------:|----------:|
| 'SpreadSheetTasks - XLSB Read - v1' | 200kFile.xlsb        |  94.33 ms | 1.756 ms | 2.462 ms | 5333.3333 | 3833.3333 | 1333.3333 |  68.48 MB |
| 'SpreadSheetTasks - XLSB Read - v2' | 200kFile.xlsb        | 114.44 ms | 2.272 ms | 4.590 ms | 5000.0000 | 3500.0000 | 1000.0000 |  49.03 MB |
| 'Sylvan.Data.Excel - XLSB Read'     | 200kFile.xlsb        | 124.67 ms | 2.353 ms | 2.311 ms | 6000.0000 | 3000.0000 | 1000.0000 |  50.82 MB |
| 'SpreadSheetTasks - XLSB Read - v1' | 65K_R(...).xlsb [21] |  48.88 ms | 0.498 ms | 0.416 ms | 2363.6364 |  818.1818 |  727.2727 |  28.83 MB |
| 'SpreadSheetTasks - XLSB Read - v2' | 65K_R(...).xlsb [21] |  63.76 ms | 0.726 ms | 0.643 ms | 1625.0000 |         - |         - |  13.66 MB |
| 'Sylvan.Data.Excel - XLSB Read'     | 65K_R(...).xlsb [21] |  72.50 ms | 0.714 ms | 0.668 ms | 2857.1429 |  142.8571 |         - |  23.16 MB |

(v1 means using UseMemoryStreamInXlsb property, v2 means using UseMemoryStreamInXlsb = false, v1 is faster but uses more memory)

### XLSB/XLSX Write

| Method                          | ReaderType | Mean     | Error   | StdDev  | Gen0      | Gen1      | Allocated |
|-------------------------------- |----------- |---------:|--------:|--------:|----------:|----------:|----------:|
| 'SpreadSheetTasks - XLSB Write' | GENERAL    | 117.7 ms | 1.49 ms | 1.24 ms | 1600.0000 |         - |  30.57 MB |
| XlsbSylvanWrite                 | GENERAL    | 163.7 ms | 3.06 ms | 3.14 ms | 1500.0000 |         - |  36.75 MB |
| 'SpreadSheetTasks - XLSX Write' | GENERAL    | 529.2 ms | 5.75 ms | 5.10 ms | 1000.0000 |         - |  30.74 MB |
| 'Sylvan - XLSX Write'           | GENERAL    | 227.4 ms | 4.01 ms | 3.75 ms | 2500.0000 | 1000.0000 |  42.94 MB |

https://github.com/KrzysztofDusko/SpreadSheetTasks
