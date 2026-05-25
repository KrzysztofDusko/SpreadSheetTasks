using SpreadSheetTasks;
using System.Data;
using System.IO.Compression;

namespace Tests;

[Collection("Sequential")]
public class MissingApiTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetRowsOfSheet_ReturnsExpectedRowCount(string extension)
    {
        var fileName = $"test_get_rows_count{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Rows.Add("A", 1);
        dt.Rows.Add("B", 2);
        dt.Rows.Add("C", 3);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            var rows = reader.GetRowsOfSheet("Sheet1").Count();
            Assert.Equal(4, rows); // header + 3 data rows
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void RowCount_AfterOpen_ReturnsValue(string extension)
    {
        var fileName = $"test_row_count_prop{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        for (int i = 0; i < 25; i++)
        {
            dt.Rows.Add($"Row{i}");
        }

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";

            var rowCount = reader.RowCount;
            Assert.True(rowCount >= 25, $"Expected >= 25, got {rowCount}");
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ResultsCount_AfterOpen_ReturnsSheetCount(string extension)
    {
        var fileName = $"test_results_count{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
            writer.AddSheet("Sheet2");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            Assert.Equal(2, reader.ResultsCount);
        }

        File.Delete(fileName);
    }

    [Fact]
    public void ReplaceSheetData_Xlsx_ReturnsNonEmptyRange()
    {
        var fileName = "test_replace_sheet_xlsx.xlsx";

        DataTable originalDt = new DataTable();
        originalDt.Columns.Add("Col1", typeof(string));
        originalDt.Columns.Add("Col2", typeof(int));
        originalDt.Rows.Add("Old1", 100);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(originalDt.CreateDataReader());
        }

        DataTable newDt = new DataTable();
        newDt.Columns.Add("Col1", typeof(string));
        newDt.Columns.Add("Col2", typeof(int));
        newDt.Rows.Add("New1", 300);

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName, updateMode: true);
            var range = reader.ReplaceSheetData("Sheet1", newDt.CreateDataReader());
            Assert.NotNull(range);
            Assert.NotEmpty(range);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(4096, true, true)]
    [InlineData(8192, false, true)]
    [InlineData(4096, true, false)]
    public void XlsxWriter_AdvancedConstructor_VariousParams_Works(int bufferSize, bool inMemoryMode, bool useSharedStrings)
    {
        var fileName = $"test_xlsx_adv_{bufferSize}_{inMemoryMode}_{useSharedStrings}.xlsx";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = new XlsxWriter(fileName, bufferSize, inMemoryMode, useSharedStrings))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read());
            Assert.Equal("Data", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(CompressionLevel.Fastest)]
    [InlineData(CompressionLevel.Optimal)]
    [InlineData(CompressionLevel.NoCompression)]
    public void XlsbWriter_AdvancedConstructor_VariousCompression_Works(CompressionLevel compressionLevel)
    {
        var fileName = $"test_xlsb_comp_{compressionLevel}.xlsb";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = new XlsbWriter(fileName, compressionLevel))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read());
            Assert.Equal("Data", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Fact]
    public void XlsbWriter_WriteSheet_StringArray_Works()
    {
        var fileName = "test_xlsb_string_array.xlsb";

        string[] data = new string[] { "Alpha", "Beta", "Gamma" };

        using (var writer = new XlsbWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(data);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read()); // first row is data (no header with string[])
            Assert.Equal("Alpha", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Beta", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Gamma", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Fact]
    public void XlsxWriter_WriteSheet_DataTableOverride_Works()
    {
        var fileName = "test_xlsx_datatable_override.xlsx";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Rows.Add("A", 1);
        dt.Rows.Add("B", 2);

        using (var writer = new XlsxWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt, headers: true);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read()); // header
            Assert.Equal("Col1", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("A", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Fact]
    public void XlsxWriter_TryToSpecifyWidthForMemoryMode_CanBeSet()
    {
        var fileName = "test_specify_width.xlsx";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = new XlsxWriter(fileName))
        {
            writer.TryToSpecifyWidthForMemoryMode = true;
            Assert.True(writer.TryToSpecifyWidthForMemoryMode);

            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read());
            Assert.Equal("Data", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Fact]
    public void XlsbWriter_Constructor_WithStream_Works()
    {
        using var memoryStream = new MemoryStream();

        using (var writer = new XlsbWriter(memoryStream, CompressionLevel.Optimal, leaveExcelArchiveOpen: true))
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            dt.Rows.Add("Data");

            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        // Stream should have data written to it when leaveExcelArchiveOpen is true
        Assert.True(memoryStream.Length > 0);
    }

    [Fact]
    public void XlsxWriter_Constructor_WithStream_Works()
    {
        using var memoryStream = new MemoryStream();

        using (var writer = new XlsxWriter(memoryStream, 4096, true, true, CompressionLevel.Optimal, leaveExcelArchiveOpen: true))
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            dt.Rows.Add("Data");

            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(memoryStream.Length > 0);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_ReadsBackCorrectData(string extension)
    {
        var fileName = $"test_readback{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Rows.Add("Row1", 1);
        dt.Rows.Add("Row2", 2);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read()); // header
            Assert.True(reader.Read());
            Assert.Equal("Row1", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Row2", reader.GetValue(0));
        }

        File.Delete(fileName);
    }
}
