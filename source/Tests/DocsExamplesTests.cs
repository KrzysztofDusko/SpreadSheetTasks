using SpreadSheetTasks;
using System.Data;
using System.Text;

namespace Tests;

[Collection("Sequential")]
public class DocsExamplesTests
{
    private static string GetTempPath(string extension)
    {
        return $"docs_test_{Guid.NewGuid():N}{extension}";
    }

    [Fact]
    public void Example_WriteFromDataTable_Xlsx()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Salary", typeof(double));
            dt.Rows.Add("Alice", 30, 5000.0);
            dt.Rows.Add("Bob", 25, 4500.0);

            using (var writer = new XlsxWriter(path))
            {
                writer.AddSheet("Employees");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Employees";
                Assert.True(reader.Read()); // header
                Assert.True(reader.Read()); // data row 1
                Assert.Equal("Alice", reader.GetValue(0));
                Assert.Equal(30L, reader.GetValue(1));
                Assert.True(reader.Read()); // data row 2
                Assert.Equal("Bob", reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_WriteFromDataTable_Xlsb()
    {
        var path = GetTempPath(".xlsb");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Salary", typeof(double));
            dt.Rows.Add("Alice", 30, 5000.0);
            dt.Rows.Add("Bob", 25, 4500.0);

            using (var writer = new XlsbWriter(path))
            {
                writer.AddSheet("Employees");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Employees";
                Assert.True(reader.Read());
                Assert.True(reader.Read());
                Assert.Equal("Alice", reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_WriteFromList()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var headers = new List<string> { "Product", "Price", "Quantity" };
            var types = new List<TypeCode> { TypeCode.String, TypeCode.Double, TypeCode.Int32 };
            var rows = new List<object?[]>
            {
                new object?[] { "Apple", 1.99, 100 },
                new object?[] { "Banana", 0.99, 250 },
                new object?[] { "Cherry", 3.49, 75 },
            };

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Products");
                writer.WriteSheet(headers, types, rows, headers: true, doAutofilter: true);
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Products";
                Assert.True(reader.Read()); // header
                Assert.Equal("Product", reader.GetValue(0));
                Assert.True(reader.Read()); // data row 1
                Assert.Equal("Apple", reader.GetString(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_WriteFromStringArray()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            string[] data = ["Alpha", "Beta", "Gamma"];

            using (var writer = new XlsxWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(data);
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                Assert.Equal("Alpha", reader.GetValue(0));
                Assert.True(reader.Read());
                Assert.Equal("Beta", reader.GetValue(0));
                Assert.True(reader.Read());
                Assert.Equal("Gamma", reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_MultiSheetWrite()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Value", typeof(int));
            dt.Rows.Add(1);
            dt.Rows.Add(2);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());

                writer.AddSheet("Sheet2");
                writer.WriteSheet(dt.CreateDataReader());

                writer.AddSheet("HiddenSheet", hidden: true);
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                Assert.Equal(3, reader.GetScheetNames().Length);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                Assert.True(reader.Read());
                Assert.Equal(1L, reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_WriteWithAutofilter()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("City", typeof(string));
            dt.Columns.Add("Population", typeof(int));
            dt.Rows.Add("New York", 8_400_000);
            dt.Rows.Add("Los Angeles", 3_800_000);
            dt.Rows.Add("Chicago", 2_700_000);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Cities");
                writer.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Cities";
                Assert.True(reader.Read()); // header
                Assert.True(reader.Read()); // data
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_WriteWithFormattedCells()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Value", typeof(object));

            dt.Rows.Add("Thousands", new FormattedCell(1234567, F.THOUSANDS_SEP));
            dt.Rows.Add("Currency PLN", new FormattedCell(1234.56, F.CURRENCY_PLN));
            dt.Rows.Add("Date", new FormattedCell(new DateTime(2026, 6, 1), F.DATE_SHORT));
            dt.Rows.Add("DateTime ISO", new FormattedCell(new DateTime(2026, 6, 1, 14, 30, 0), F.DATETIME_ISO));
            dt.Rows.Add("Percentage", new FormattedCell(0.25, F.PERCENTAGE));
            dt.Rows.Add("Scientific", new FormattedCell(12345.67, F.SCIENTIFIC));
            dt.Rows.Add("Time", new FormattedCell(new DateTime(2026, 6, 1, 8, 15, 0), F.TIME_HH_MM_SS));

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Formatted");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Formatted";
                Assert.True(reader.Read()); // header
                Assert.True(reader.Read()); // first data row
                Assert.Equal("Thousands", reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_WriteWithFormattedCells_CustomFormat()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Desc", typeof(string));
            dt.Columns.Add("Value", typeof(object));
            dt.Rows.Add("Custom", new FormattedCell(1234.56, "#,##0.00 \"USD\""));
            dt.Rows.Add("ZIP", new FormattedCell(12345, "00000"));

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                Assert.True(reader.Read());
                Assert.Equal("Custom", reader.GetValue(0));
                Assert.True(reader.Read());
                Assert.Equal("ZIP", reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_WriteToFileStream()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            using (var fs = File.Open(path, FileMode.Create))
            using (var writer = new XlsxWriter(fs))
            {
                writer.AddSheet("Sheet1");
                var dt = new DataTable();
                dt.Columns.Add("Col1", typeof(string));
                dt.Rows.Add("Hello");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                Assert.True(reader.Read());
                Assert.Equal("Hello", reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_WriteToMemoryStream()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            byte[] excelBytes;
            using (var ms = new MemoryStream())
            {
                using (var writer = new XlsxWriter(ms))
                {
                    writer.AddSheet("Sheet1");
                    var dt = new DataTable();
                    dt.Columns.Add("Col1", typeof(string));
                    dt.Rows.Add("World");
                    writer.WriteSheet(dt.CreateDataReader());
                }
                excelBytes = ms.ToArray();
            }
            Assert.True(excelBytes.Length > 0);

            File.WriteAllBytes(path, excelBytes);
            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                Assert.True(reader.Read());
                Assert.Equal("World", reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_HiddenSheets()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Data", typeof(string));
            dt.Rows.Add("visible");
            dt.Rows.Add("hidden data");

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Visible");
                writer.WriteSheet(dt.CreateDataReader());

                writer.AddSheet("Hidden", hidden: true);
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                Assert.Equal(2, reader.GetScheetNames().Length);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_Events_OnCompress_On10k()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Index", typeof(int));
            for (int i = 0; i < 25_000; i++)
                dt.Rows.Add(i);

            int compressFired = 0;
            int tenKFired = 0;

            using (var writer = new XlsxWriter(path))
            {
                writer.OnCompress += () => compressFired++;
                writer.On10k += (row) => tenKFired++;

                writer.AddSheet("LargeData");
                writer.WriteSheet(dt.CreateDataReader());
            }

            Assert.Equal(1, compressFired);
            Assert.True(tenKFired >= 2); // 25k rows = 2x On10k
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_DocumentProperties()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            dt.Rows.Add("Data");

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.DocPopertyProgramName = "MyApplication";
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            Assert.True(File.Exists(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_SuppressSomeDate()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Rows.Add(new DateTime(1000, 1, 1));
            dt.Rows.Add(new DateTime(2026, 6, 1));

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.SuppressSomeDate = true;
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            Assert.True(File.Exists(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_BasicRead()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("Alice", 30);
            dt.Rows.Add("Bob", 25);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            var sb = new StringBuilder();
            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";

                object[]? row = null;
                while (reader.Read())
                {
                    row ??= new object[reader.FieldCount];
                    reader.GetValues(row);
                    sb.AppendLine(string.Join("|", row));
                }
            }

            Assert.Contains("Alice", sb.ToString());
            Assert.Contains("Bob", sb.ToString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_ReadWithTypedGetters()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Salary", typeof(double));
            dt.Rows.Add("Alice", 30, 5000.0);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";

                while (reader.Read())
                {
                    reader.GetString(0);  // Name
                }
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_ReadWithTypedGetters_MixedTypes()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Salary", typeof(double));
            dt.Rows.Add("Alice", 30, 5000.0);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader(), headers: false);
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                Assert.Equal("Alice", reader.GetString(0));
                Assert.Equal(30, reader.GetInt32(1));
                Assert.Equal(5000.0, reader.GetDouble(2), 3);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_GetRowsOfSheet()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            dt.Columns.Add("Col2", typeof(int));
            dt.Rows.Add("A", 1);
            dt.Rows.Add("B", 2);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                int totalRows = reader.GetRowsOfSheet("Sheet1").Count();
                Assert.Equal(3, totalRows); // header + 2 data rows
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_GetSheetNames()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            dt.Rows.Add("Data");

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("First");
                writer.WriteSheet(dt.CreateDataReader());
                writer.AddSheet("Second");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                string[] sheetNames = reader.GetScheetNames();
                Assert.Equal(2, sheetNames.Length);
                Assert.Contains("First", sheetNames);
                Assert.Contains("Second", sheetNames);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_RowCount_ResultsCount()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            for (int i = 0; i < 10; i++)
                dt.Rows.Add($"Row{i}");

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                int rowCount = reader.RowCount;
                int sheetCount = reader.ResultsCount;
                Assert.True(rowCount >= 10);
                Assert.Equal(1, sheetCount);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_TreatAllColumnsAsText()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("Alice", 30);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader(), headers: false);
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.TreatAllColumnsAsText = true;
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                string val = reader.GetString(1);
                Assert.Equal("30", val);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_UseMemoryStreamInXlsb()
    {
        var path = GetTempPath(".xlsb");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            dt.Rows.Add("Test");

            using (var writer = new XlsbWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.UseMemoryStreamInXlsb = false;
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_UpdateMode_ReplaceSheetData()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            var original = new DataTable();
            original.Columns.Add("Name", typeof(string));
            original.Columns.Add("Score", typeof(int));
            original.Rows.Add("Old", 0);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(original.CreateDataReader());
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path, updateMode: true);

                var newData = new DataTable();
                newData.Columns.Add("Name", typeof(string));
                newData.Columns.Add("Score", typeof(int));
                newData.Rows.Add("Alice", 95);
                newData.Rows.Add("Bob", 87);

                string range = reader.ReplaceSheetData("Sheet1", newData.CreateDataReader());
                Assert.NotNull(range);
                Assert.NotEmpty(range);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Example_FactoryMethod()
    {
        var pathXlsx = GetTempPath(".xlsx");
        var pathXlsb = GetTempPath(".xlsb");
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            dt.Rows.Add("Hello");

            using (var writer = ExcelWriter.CreateWriter(pathXlsx))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            using (var writer = ExcelWriter.CreateWriter(pathXlsb))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(dt.CreateDataReader());
            }

            Assert.True(File.Exists(pathXlsx));
            Assert.True(File.Exists(pathXlsb));
        }
        finally
        {
            if (File.Exists(pathXlsx)) File.Delete(pathXlsx);
            if (File.Exists(pathXlsb)) File.Delete(pathXlsb);
        }
    }

    [Fact]
    public void Example_FormatConstants()
    {
        Assert.Equal("#,##0", F.THOUSANDS_SEP);
        Assert.Equal("#,##0.00 \"z\u0142\"", F.CURRENCY_PLN);
        Assert.Equal("#,##0.00 \u20AC", F.CURRENCY_EUR);
        Assert.Equal("0%", F.PERCENTAGE);
        Assert.Equal("0.00E+00", F.SCIENTIFIC);
        Assert.Equal("#,##0.00", F.TWO_DECIMALS);
        Assert.Equal("@", F.TEXT);
        Assert.Equal("000000000", F.LEADING_ZEROS);
        Assert.Equal("dd.mm.yyyy", F.DATE_SHORT);
        Assert.Equal("yyyy-mm-dd", F.DATE_ISO);
        Assert.Equal("hh:mm", F.TIME_HH_MM);
        Assert.Equal("hh:mm:ss", F.TIME_HH_MM_SS);
        Assert.Equal("h:mm AM/PM", F.TIME_12H);
        Assert.Equal("hh:mm:ss.000", F.TIME_MS);
    }

    [Fact]
    public void Example_FormattedCell_Construction()
    {
        var cell = new FormattedCell(42, F.THOUSANDS_SEP);
        Assert.Equal(42, cell.Value);
        Assert.Equal(F.THOUSANDS_SEP, cell.Format);
    }

    [Fact]
    public void Example_FormattedCellFromList()
    {
        var path = GetTempPath(".xlsx");
        try
        {
            using (var writer = ExcelWriter.CreateWriter(path))
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

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                Assert.True(reader.Read());
                Assert.Equal("Thousands", reader.GetValue(0));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
