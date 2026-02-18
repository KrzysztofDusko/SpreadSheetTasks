using SpreadSheetTasks;
using System.Data;

DataTable dt = new DataTable();
dt.Columns.Add("Col1", typeof(int));
dt.Columns.Add("Col2", typeof(string));
dt.Columns.Add("Col3", typeof(int));
dt.Rows.Add(1, DBNull.Value, 2);

using (var writer = ExcelWriter.CreateWriter("test_null_debug.xlsx"))
{
    writer.AddSheet("Sheet1");
    writer.WriteSheet(dt.CreateDataReader());
}

Console.WriteLine("File created: test_null_debug.xlsx");

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("test_null_debug.xlsx");
    reader.ActualSheetName = "Sheet1";
    
    reader.Read(); // Skip header
    reader.Read();
    
    Console.WriteLine($"Col0: {reader.GetValue(0)} (type: {reader.GetValue(0)?.GetType().Name ?? "null"})");
    Console.WriteLine($"Col1: {reader.GetValue(1)} (type: {reader.GetValue(1)?.GetType().Name ?? "null"})");
    Console.WriteLine($"Col2: {reader.GetValue(2)} (type: {reader.GetValue(2)?.GetType().Name ?? "null"})");
}
