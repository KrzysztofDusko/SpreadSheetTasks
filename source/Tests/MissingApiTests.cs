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
            writer.WriteSheet(dt);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            Assert.True(reader.Read());

            var rowCount = reader.RowCount;
            if (extension == ".xlsb")
                Assert.Equal(-1, rowCount);
            else
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

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithFormattedCell_CreatesFile(string extension)
    {
        var fileName = $"test_formatted_cell{extension}";

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            var dt = new DataTable();
            dt.Columns.Add("Desc", typeof(string));
            dt.Columns.Add("Value", typeof(object));
            dt.Rows.Add("Number", new FormattedCell(1234567, F.THOUSANDS_SEP));
            dt.Rows.Add("Currency", new FormattedCell(1234.56, F.CURRENCY_PLN));
            dt.Rows.Add("Date", new FormattedCell(new DateTime(2026, 6, 1), F.DATE_SHORT));
            dt.Rows.Add("DateTime", new FormattedCell(new DateTime(2026, 6, 1, 14, 34, 0), F.DATETIME_ISO));
            dt.Rows.Add("Pct", new FormattedCell(0.25, F.PERCENTAGE));
            dt.Rows.Add("Sci", new FormattedCell(12345.67, F.SCIENTIFIC));
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(File.Exists(fileName));

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read()); // first data row
            Assert.Equal("Number", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithFormattedCellFromList_CreatesFile(string extension)
    {
        var fileName = $"test_formatted_list{extension}";

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            var headers = new List<string> { "Desc", "Value" };
            var types = new List<TypeCode> { TypeCode.String, TypeCode.Object };
            var rows = new List<object?[]>
            {
                new object?[] { "Thousands", new FormattedCell(1234567, F.THOUSANDS_SEP) },
                new object?[] { "Date ISO", new FormattedCell(new DateTime(2026, 6, 1), F.DATE_ISO) },
                new object?[] { "Custom", new FormattedCell(1234.56, "#,##0.00 \"USD\"") },
                new object?[] { "Time", new FormattedCell(new DateTime(2026, 6, 1, 14, 34, 0), F.TIME_HH_MM) },
            };
            writer.WriteSheet(headers, types, rows, headers: true);
        }

        Assert.True(File.Exists(fileName));

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read()); // first data row
            Assert.Equal("Thousands", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Fact]
    public void FormattedCell_Constructed_CreatesValidInstance()
    {
        var cell = new FormattedCell(42, F.THOUSANDS_SEP);
        Assert.Equal(42, cell.Value);
        Assert.Equal(F.THOUSANDS_SEP, cell.Format);
    }

    [Fact]
    public void F_FormatConstants_AllDefined()
    {
        Assert.Equal("#,##0", F.THOUSANDS_SEP);
        Assert.Equal("dd.mm.yyyy", F.DATE_SHORT);
        Assert.Equal("yyyy-mm-dd\"T\"hh:mm:ss", F.DATETIME_ISO);
        Assert.Equal("hh:mm", F.TIME_HH_MM);
    }

    [Fact]
    public void ReplacePivotTableDim_NoPivotTable_ThrowsExpectedException()
    {
        var fileName = "test_no_pivot.xlsx";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName, updateMode: true);
            var range = reader.ReplaceSheetData("Sheet1", dt.CreateDataReader());

            var ex = Assert.Throws<Exception>(() => reader.ReplacePivotTableDim("NonExistentPivot", range));
            Assert.Contains("not found", ex.Message);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetRowsOfSheet_MultipleSheets_ReturnsCorrectRowCounts(string extension)
    {
        var fileName = $"test_get_rows_multi{extension}";

        DataTable dt1 = new DataTable();
        dt1.Columns.Add("Col1", typeof(string));
        dt1.Rows.Add("A1");
        dt1.Rows.Add("A2");

        DataTable dt2 = new DataTable();
        dt2.Columns.Add("Col1", typeof(string));
        dt2.Rows.Add("B1");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt1.CreateDataReader());
            writer.AddSheet("Sheet2");
            writer.WriteSheet(dt2.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);

            var rows1 = reader.GetRowsOfSheet("Sheet1").Count();
            Assert.Equal(3, rows1); // header + 2 data rows

            var rows2 = reader.GetRowsOfSheet("Sheet2").Count();
            Assert.Equal(2, rows2); // header + 1 data row
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetRowsOfSheet_InvalidSheetName_ThrowsException(string extension)
    {
        var fileName = $"test_get_rows_invalid{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            Assert.ThrowsAny<Exception>(() => reader.GetRowsOfSheet("NonExistent").Count());
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_ListOverload_ReadsBackCorrectData(string extension)
    {
        var fileName = $"test_list_readback{extension}";

        var headers = new List<string> { "Name", "Age", "Salary" };
        var types = new List<TypeCode> { TypeCode.String, TypeCode.Int32, TypeCode.Double };
        var rows = new List<object?[]>
        {
            new object?[] { "Alice", 30, 5000.0 },
            new object?[] { "Bob", 25, 4500.0 }
        };

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(headers, types, rows, headers: true);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read()); // header
            Assert.Equal("Name", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Alice", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Bob", reader.GetValue(0));
        }

        File.Delete(fileName);
    }
}
