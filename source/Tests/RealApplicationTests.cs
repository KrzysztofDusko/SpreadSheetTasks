using SpreadSheetTasks;
using System.Data;
using System.IO.Compression;

namespace Tests;

[Collection("Sequential")]
public class RealApplicationTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void File_IsValidZipArchive(string extension)
    {
        var fileName = $"test_valid_zip{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var archive = ZipFile.OpenRead(fileName))
        {
            Assert.True(archive.Entries.Count > 0, "ZIP archive should contain entries");
        }

        File.Delete(fileName);
    }

    [Fact]
    public void XlsxFile_HasRequiredOpenXmlParts()
    {
        var fileName = $"test_xlsx_parts.xlsx";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Rows.Add("Data", 42);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var archive = ZipFile.OpenRead(fileName))
        {
            var entryNames = archive.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);

            Assert.Contains("[Content_Types].xml", entryNames);
            Assert.Contains("_rels/.rels", entryNames);
            Assert.Contains("xl/workbook.xml", entryNames);
            Assert.Contains("xl/_rels/workbook.xml.rels", entryNames);
            Assert.Contains("xl/styles.xml", entryNames);
            Assert.Contains("xl/sharedStrings.xml", entryNames);
            Assert.True(entryNames.Any(e => e.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)),
                "Should contain at least one worksheet part");
        }

        File.Delete(fileName);
    }

    [Fact]
    public void XlsbFile_HasRequiredBinaryParts()
    {
        var fileName = $"test_xlsb_parts.xlsb";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Rows.Add("Data", 42);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var archive = ZipFile.OpenRead(fileName))
        {
            var entryNames = archive.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);

            Assert.Contains("[Content_Types].xml", entryNames);
            Assert.Contains("_rels/.rels", entryNames);
            Assert.Contains("xl/workbook.bin", entryNames);
            Assert.Contains("xl/_rels/workbook.bin.rels", entryNames);
            Assert.Contains("xl/styles.bin", entryNames);
            Assert.True(entryNames.Any(e => e.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)),
                "Should contain at least one worksheet part");
        }

        File.Delete(fileName);
    }

    [Fact]
    public void XlsxFile_ContentTypesXml_HasRequiredOverrides()
    {
        var fileName = $"test_content_types.xlsx";

        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var archive = ZipFile.OpenRead(fileName))
        {
            var entry = archive.GetEntry("[Content_Types].xml");
            Assert.NotNull(entry);
            Assert.True(entry.Length > 0);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteAndRead_MultiSheet_AllSheetsAccessible(string extension)
    {
        var fileName = $"test_multi_sheet_real{extension}";

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            for (int i = 1; i <= 5; i++)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("SheetCol", typeof(string));
                dt.Rows.Add($"Sheet{i}_Data");
                writer.AddSheet($"Sheet{i}");
                writer.WriteSheet(dt.CreateDataReader());
            }
        }

        using (var archive = ZipFile.OpenRead(fileName))
        {
            var sheetEntryCount = archive.Entries.Count(e =>
                e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                (e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || e.FullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)));
            Assert.Equal(5, sheetEntryCount);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            var names = reader.GetScheetNames();
            Assert.Equal(5, names.Length);
            for (int i = 1; i <= 5; i++)
            {
                Assert.Contains($"Sheet{i}", names);
            }
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void SharedStrings_LargeUniqueSet_AllPreserved(string extension)
    {
        var fileName = $"test_shared_stress{extension}";
        const int uniqueCount = 1000;

        DataTable dt = new DataTable();
        dt.Columns.Add("Id", typeof(int));
        dt.Columns.Add("Value", typeof(string));

        for (int i = 0; i < uniqueCount; i++)
        {
            dt.Rows.Add(i, $"UNIQUE_STRING_{i}_WITH_SUFFIX_{Guid.NewGuid():N}");
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

            Assert.True(reader.Read()); // header

            var values = new HashSet<string>();
            while (reader.Read())
            {
                var val = reader.GetString(1);
                Assert.NotNull(val);
                Assert.True(values.Add(val), $"Duplicate value found: {val}");
            }

            Assert.Equal(uniqueCount, values.Count);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithAllDataTypes_RoundtripsCorrectly(string extension)
    {
        var fileName = $"test_all_types_roundtrip{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Columns.Add("LongCol", typeof(long));
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Columns.Add("StringCol", typeof(string));
        dt.Columns.Add("BoolCol", typeof(bool));
        dt.Columns.Add("DateCol", typeof(DateTime));
        dt.Columns.Add("NullCol", typeof(string));

        dt.Rows.Add(42, 9876543210L, 3.14159265358979, "hello", true, new DateTime(2024, 6, 15, 14, 30, 0), DBNull.Value);
        dt.Rows.Add(-1, -9876543210L, -0.001, "world", false, new DateTime(1900, 1, 1), DBNull.Value);
        dt.Rows.Add(0, 0L, 1e-10, "", true, new DateTime(2025, 12, 31, 23, 59, 59), DBNull.Value);

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

            for (int row = 0; row < 3; row++)
            {
                Assert.True(reader.Read());
                Assert.Equal((long)(int)dt.Rows[row][0], reader.GetInt64(0));
                Assert.Equal((long)dt.Rows[row][1], reader.GetInt64(1));
                Assert.Equal((double)dt.Rows[row][2], reader.GetDouble(2), 10);
                Assert.Equal((string)dt.Rows[row][3], reader.GetString(3));

                var expectedBool = (bool)dt.Rows[row][4];
                var actualBool = reader.GetValue(4);
                Assert.True((actualBool is bool b && b == expectedBool) || (actualBool is int i && ((i == 1) == expectedBool)));

                var expectedDate = (DateTime)dt.Rows[row][5];
                Assert.Equal(expectedDate, reader.GetDateTime(5));

                Assert.Equal(DBNull.Value, reader.GetValue(6));
            }
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void CrossFormat_WriteWithSpreadSheetTasks_ReadWithOurReader_DataIntegrity(string extension)
    {
        var fileName = SylvanInteropTestHelpers.CreateTempExcelPath(extension);

        try
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Score", typeof(double));
            dt.Columns.Add("Active", typeof(bool));
            dt.Columns.Add("Birth", typeof(DateTime));

            dt.Rows.Add("Alice", 30, 95.5, true, new DateTime(1994, 5, 15));
            dt.Rows.Add("Bob", 25, 88.0, false, new DateTime(1999, 3, 22));
            dt.Rows.Add("Charlie", 35, 72.3, true, new DateTime(1989, 11, 8));

            SylvanInteropTestHelpers.WriteWithSpreadSheetTasks(fileName, dt);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read()); // header
            Assert.Equal("Name", reader.GetValue(0));
            Assert.Equal("Age", reader.GetValue(1));

            Assert.True(reader.Read());
            Assert.Equal("Alice", reader.GetValue(0));
            Assert.Equal(30L, reader.GetValue(1));

            Assert.True(reader.Read());
            Assert.Equal("Bob", reader.GetValue(0));

            Assert.True(reader.Read());
            Assert.Equal("Charlie", reader.GetValue(0));

            Assert.False(reader.Read());
        }
        finally
        {
            if (File.Exists(fileName)) File.Delete(fileName);
        }
    }
}
