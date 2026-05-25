using SpreadSheetTasks;
using Sylvan.Data.Excel;
using System.Data;
using System.Globalization;

namespace Tests;

[Collection("Sequential")]
public class FixesVerificationTests
{
    private const string SheetName = "Sheet1";

    /// <summary>
    /// Verifies Fix #1: _numberFormatsTypeDictionary is per-instance, not static.
    /// Two independent reader instances should not share format type state.
    /// This writes two files with our own writer (predictable format IDs),
    /// then reads each with a separate reader instance.
    /// </summary>
    [Theory]
    [InlineData(".xlsb")]
    [InlineData(".xlsx")]
    public void NumberFormatTypeDictionary_IsPerInstance_NoCrossContamination(string extension)
    {
        string file1 = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        string file2 = SylvanInteropTestHelpers.CreateTempExcelPath(extension);

        try
        {
            // File 1: DateTime column + int column
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("DateCol", typeof(DateTime));
            dt1.Columns.Add("IntCol", typeof(int));
            dt1.Rows.Add(new DateTime(2024, 6, 15), 42);

            // File 2: String column + int column (same structure, different types)
            DataTable dt2 = new DataTable();
            dt2.Columns.Add("TextCol", typeof(string));
            dt2.Columns.Add("IntCol", typeof(int));
            dt2.Rows.Add("hello world", 99);

            // Use our own writer for predictable output
            SylvanInteropTestHelpers.WriteWithSpreadSheetTasks(file1, dt1);
            SylvanInteropTestHelpers.WriteWithSpreadSheetTasks(file2, dt2);

            // Read with independent reader instances
            var reader1 = new XlsxOrXlsbReadOrEdit();
            reader1.Open(file1);
            reader1.ActualSheetName = SheetName;
            Assert.True(reader1.Read()); // header
            Assert.True(reader1.Read()); // first data row
            var val1 = reader1.GetValue(0); // DateCol
            reader1.Dispose();

            var reader2 = new XlsxOrXlsbReadOrEdit();
            reader2.Open(file2);
            reader2.ActualSheetName = SheetName;
            Assert.True(reader2.Read()); // header
            Assert.True(reader2.Read()); // first data row
            var val2 = reader2.GetValue(0); // TextCol
            reader2.Dispose();

            // File1: DateTime column — our writer uses format ID 3 (DateTime)
            // which maps to typeof(DateTime?) in _numberFormatsTypeDictionary
            Assert.Equal(new DateTime(2024, 6, 15), val1);

            // File2: String column — gets written as shared string,
            // NOT as a numeric value with custom format
            Assert.Equal("hello world", val2);
        }
        finally
        {
            if (File.Exists(file1)) File.Delete(file1);
            if (File.Exists(file2)) File.Delete(file2);
        }
    }

    /// <summary>
    /// Verifies Fix #2: RK double values in XLSB files are read correctly.
    /// Uses our own writer (full IEEE 754, type 0x05) and Sylvan writer as reference.
    /// Both must produce the same values when read by our reader.
    /// </summary>
    [Theory]
    [InlineData(".xlsb")]
    [InlineData(".xlsx")]
    public void CrossFormat_DoubleValues_ReadBackCorrectly(string extension)
    {
        string spreadPath = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        string sylvanPath = SylvanInteropTestHelpers.CreateTempExcelPath(extension);

        try
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DoubleCol", typeof(double));
            dt.Rows.Add(3.14159);
            dt.Rows.Add(-0.001);
            dt.Rows.Add(1e10);
            dt.Rows.Add(0.0);
            dt.Rows.Add(42.0);

            SylvanInteropTestHelpers.WriteWithSpreadSheetTasks(spreadPath, dt);
            SylvanInteropTestHelpers.WriteWithSylvan(sylvanPath, dt);

            var spreadReadSpread = ReadDataRowsWithSpreadSheetTasksRaw(spreadPath);
            var spreadReadSylvan = ReadDataRowsWithSpreadSheetTasksRaw(sylvanPath);

            // Remove headers
            var spreadData = spreadReadSpread.Skip(1).ToList();
            var sylvanData = spreadReadSylvan.Skip(1).ToList();

            Assert.Equal(spreadData.Count, sylvanData.Count);
            for (int i = 0; i < spreadData.Count; i++)
            {
                Assert.Equal(
                    Convert.ToDecimal(spreadData[i][0], CultureInfo.InvariantCulture),
                    Convert.ToDecimal(sylvanData[i][0], CultureInfo.InvariantCulture),
                    10);
            }
        }
        finally
        {
            if (File.Exists(spreadPath)) File.Delete(spreadPath);
            if (File.Exists(sylvanPath)) File.Delete(sylvanPath);
        }
    }

    /// <summary>
    /// Verifies that Fix #2 enables correct reading of RK doubles from XLSB files
    /// created by Sylvan. Direct value comparison between SpreadSheetTasks and Sylvan readers.
    /// </summary>
    [Fact]
    public void SylvanXlsb_RkDoubles_MatchSylvanReader()
    {
        string file = SylvanInteropTestHelpers.CreateTempExcelPath(".xlsb");

        try
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DoubleCol", typeof(double));
            dt.Rows.Add(1.5);
            dt.Rows.Add(-3.25);
            dt.Rows.Add(0.1);
            dt.Rows.Add(123456.789);
            dt.Rows.Add(0.0);

            SylvanInteropTestHelpers.WriteWithSylvan(file, dt);

            // Read SpreadSheetTasks reader output (includes init header row)
            var spreadRows = ReadDataRowsWithSpreadSheetTasksRaw(file);
            // Read Sylvan reader output as reference (no header row)
            var sylvanRows = ReadDataRowsWithSylvanRaw(file);

            // SpreadSheetTasks adds an initialization row; remove it for comparison
            var spreadData = spreadRows.Skip(1).ToList();
            // Sylvan reader does not include a header row for DataTableReader data
            var sylvanData = sylvanRows;

            Assert.Equal(spreadData.Count, sylvanData.Count);
            for (int i = 0; i < spreadData.Count; i++)
            {
                Assert.Equal(
                    Convert.ToDecimal(sylvanData[i][0], CultureInfo.InvariantCulture),
                    Convert.ToDecimal(spreadData[i][0], CultureInfo.InvariantCulture),
                    10);
            }
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
        }
    }

    private static List<object?[]> ReadDataRowsWithSpreadSheetTasksRaw(string path)
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(path);
        reader.ActualSheetName = SheetName;

        var rows = new List<object?[]>();
        while (reader.Read())
        {
            object[] row = new object[reader.FieldCount];
            reader.GetValues(row);
            rows.Add(row);
        }

        return rows;
    }

    private static List<object?[]> ReadDataRowsWithSylvanRaw(string path)
    {
        using var reader = ExcelDataReader.Create(path);

        var rows = new List<object?[]>();
        while (reader.Read())
        {
            object[] row = new object[reader.FieldCount];
            reader.GetValues(row);
            rows.Add(row);
        }

        return rows;
    }
}
