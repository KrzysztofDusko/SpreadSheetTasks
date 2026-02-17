using SpreadSheetTasks;
using System.Data;

namespace Tests;

[Collection("Sequential")]
public class WriterTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void CreateWriter_ValidExtension_CreatesWriter(string extension)
    {
        var fileName = $"test_create{extension}";
        
        var writer = ExcelWriter.CreateWriter(fileName);
        Assert.NotNull(writer);
        Assert.IsAssignableFrom<ExcelWriter>(writer);
        
        writer.Dispose();
        File.Delete(fileName);
    }

    [Fact]
    public void CreateWriter_InvalidExtension_ThrowsException()
    {
        Assert.Throws<Exception>(() => ExcelWriter.CreateWriter("test.txt"));
        Assert.Throws<Exception>(() => ExcelWriter.CreateWriter("test.csv"));
        Assert.Throws<Exception>(() => ExcelWriter.CreateWriter("test.pdf"));
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void RowsCount_AfterWrite_ReturnsCorrectValue(string extension)
    {
        var fileName = $"test_rowscount{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        for (int i = 0; i < 50; i++)
        {
            dt.Rows.Add($"Row{i}");
        }

        int rowsWritten;
        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
            rowsWritten = writer.RowsCount;
        }

        Assert.True(rowsWritten > 0);
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void DocPopertyProgramName_SetValue_WritesToFile(string extension)
    {
        var fileName = $"test_docprop{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.DocPopertyProgramName = "MyTestApp";
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void AddSheet_MultipleSheets_IncreasesSheetCount(string extension)
    {
        var fileName = $"test_sheetcount{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
            writer.AddSheet("Sheet2");
            writer.WriteSheet(dt.CreateDataReader());
            writer.AddSheet("Sheet3");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            Assert.Equal(3, reader.GetScheetNames().Length);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_NoHeaders_WritesDataOnly(string extension)
    {
        var fileName = $"test_noheaders{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data1");
        dt.Rows.Add("Data2");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader(), headers: false);
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
            Assert.Equal(2, count);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithHeaders_WritesHeadersAndData(string extension)
    {
        var fileName = $"test_withheaders{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("ColumnA", typeof(string));
        dt.Columns.Add("ColumnB", typeof(int));
        dt.Rows.Add("Row1", 1);
        dt.Rows.Add("Row2", 2);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader(), headers: true);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            Assert.Equal("ColumnA", reader.GetValue(0));
            Assert.Equal("ColumnB", reader.GetValue(1));
            
            Assert.True(reader.Read());
            Assert.Equal("Row1", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_Autofilter_AddsFilter(string extension)
    {
        var fileName = $"test_autofilter{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
        }

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_FromDataTable_WritesCorrectly(string extension)
    {
        var fileName = $"test_datatable{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("StringCol", typeof(string));
        dt.Columns.Add("IntCol", typeof(int));
        dt.Rows.Add("Test", 42);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt, headers: true);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            Assert.Equal("StringCol", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Test", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_SingleColumnArray_WritesCorrectly(string extension)
    {
        var fileName = $"test_onecolumn{extension}";
        
        string[] data = ["Value1", "Value2", "Value3"];

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(data);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            int count = 0;
            while (reader.Read())
            {
                Assert.Equal(data[count], reader.GetValue(0));
                count++;
            }
            Assert.Equal(3, count);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_FromList_WritesCorrectly(string extension)
    {
        var fileName = $"test_fromlist{extension}";
        
        List<string> headers = new() { "Col1", "Col2" };
        List<TypeCode> typeCodes = new() { TypeCode.Int32, TypeCode.String };
        List<object?[]> data = new()
        {
            new object?[] { 1, "A" },
            new object?[] { 2, "B" },
            new object?[] { 3, "C" }
        };

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(headers, typeCodes, data, headers: true);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            Assert.Equal("Col1", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal(1L, reader.GetValue(0)); // GetValue returns long for integers
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Save_CalledExplicitly_SavesFile(string extension)
    {
        var fileName = $"test_save{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        var writer = ExcelWriter.CreateWriter(fileName);
        writer.AddSheet("Sheet1");
        writer.WriteSheet(dt.CreateDataReader());
        writer.Save();

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Dispose_SavesFile(string extension)
    {
        var fileName = $"test_dispose{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Fact]
    public void XlsxWriter_Constructor_WithStream_Works()
    {
        using var memoryStream = new MemoryStream();
        using (var writer = new XlsxWriter(memoryStream))
        {
            Assert.NotNull(writer);
        }
    }

    [Fact]
    public void XlsbWriter_Constructor_WithStream_Works()
    {
        using var memoryStream = new MemoryStream();
        using (var writer = new XlsbWriter(memoryStream))
        {
            Assert.NotNull(writer);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void SuppressSomeDate_SetTrue_Works(string extension)
    {
        var fileName = $"test_suppressdate{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("DateCol", typeof(DateTime));
        dt.Rows.Add(new DateTime(1000, 1, 1));

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.SuppressSomeDate = true;
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }
}
