using SpreadSheetTasks;
using System.Data;

namespace Tests;

[Collection("Sequential")]
public class ReaderTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetValue_ByIndex_ReturnsCorrectValue(string extension)
    {
        var fileName = $"test_getvalue{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Columns.Add("Col3", typeof(DateTime));
        dt.Rows.Add("Test", 42, new DateTime(2024, 1, 1));

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Header row
            Assert.True(reader.Read()); // Data row
            Assert.Equal("Test", reader.GetValue(0));
            Assert.Equal(42L, reader.GetValue(1)); // GetValue returns long for integers
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetFieldType_ReturnsCorrectType(string extension)
    {
        var fileName = $"test_fieldtype{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("StringCol", typeof(string));
        dt.Columns.Add("IntCol", typeof(int));
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Columns.Add("DateCol", typeof(DateTime));
        dt.Columns.Add("BoolCol", typeof(bool));
        dt.Rows.Add("Test", 42, 3.14, DateTime.Today, true);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            Assert.Equal(typeof(string), reader.GetFieldType(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetString_ConvertsToString(string extension)
    {
        var fileName = $"test_getstring{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Rows.Add(42, 3.14159);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            var str0 = reader.GetString(0);
            var str1 = reader.GetString(1);
            
            Assert.NotNull(str0);
            Assert.NotNull(str1);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void TreatAllColumnsAsText_ReturnsStrings(string extension)
    {
        var fileName = $"test_treat_as_text{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Rows.Add(42, 3.14159);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.TreatAllColumnsAsText = true;
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            Assert.Equal(typeof(string), reader.GetFieldType(0));
            Assert.Equal(typeof(string), reader.GetFieldType(1));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Read_EmptyRows_ReturnsCorrectCount(string extension)
    {
        var fileName = $"test_empty_rows{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Row1");
        dt.Rows.Add("Row2");
        dt.Rows.Add("Row3");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            int count = 0;
            while (reader.Read())
            {
                count++;
            }
            // Header + 3 data rows = 4
            Assert.Equal(4, count);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetName_ReturnsColumnName(string extension)
    {
        var fileName = $"test_getname{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("MyColumn", typeof(string));
        dt.Rows.Add("Value");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            Assert.Equal("MyColumn", reader.GetName(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetScheetNames_ReturnsAllSheetNames(string extension)
    {
        var fileName = $"test_sheet_names{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Alpha");
            writer.WriteSheet(dt.CreateDataReader());
            writer.AddSheet("Beta");
            writer.WriteSheet(dt.CreateDataReader());
            writer.AddSheet("Gamma");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            var names = reader.GetScheetNames();
            
            Assert.Equal(3, names.Length);
            Assert.Contains("Alpha", names);
            Assert.Contains("Beta", names);
            Assert.Contains("Gamma", names);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void FieldCount_ReturnsCorrectValue(string extension)
    {
        var fileName = $"test_fieldcount{extension}";
        
        DataTable dt = new DataTable();
        for (int i = 0; i < 10; i++)
        {
            dt.Columns.Add($"Col{i}", typeof(string));
        }
        dt.Rows.Add("1", "2", "3", "4", "5", "6", "7", "8", "9", "10");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            Assert.Equal(10, reader.FieldCount);
        }

        File.Delete(fileName);
    }
}
