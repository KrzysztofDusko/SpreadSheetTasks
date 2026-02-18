using SpreadSheetTasks;
using System.Data;

namespace Tests;

[Collection("Sequential")]
public class EdgeCaseTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Read_EmptySheet_ReturnsNoRows(string extension)
    {
        var fileName = $"test_empty_sheet{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        // No rows added

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            // Only header row exists
            Assert.True(reader.Read()); // Header
            Assert.False(reader.Read()); // No data rows
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_AllNullValues_Issue6_VerifyBehavior(string extension)
    {
        // Test for issue #6 - verify that all-null rows are preserved and read as DBNull.
        var fileName = $"test_issue6_null{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Columns.Add("Col3", typeof(DateTime));
        dt.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header

            Assert.True(reader.Read());
            Assert.Equal(DBNull.Value, reader.GetValue(0));
            Assert.Equal(DBNull.Value, reader.GetValue(1));
            Assert.Equal(DBNull.Value, reader.GetValue(2));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_MixedNullValues_Issue6_Fixed(string extension)
    {
        // Test for issue #6 - verify that null cells return DBNull.Value, not previous row values
        var fileName = $"test_issue6_mixed{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Columns.Add("Col3", typeof(double));
        
        dt.Rows.Add("value1", 1, 1.5);
        dt.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value); // Row with all nulls
        dt.Rows.Add("value2", 2, 2.5);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            
            // First row - has values
            Assert.True(reader.Read());
            Assert.Equal("value1", reader.GetValue(0));
            Assert.Equal(1L, (long)reader.GetValue(1));
            Assert.Equal(1.5, (double)reader.GetValue(2));
            
            // Second row - all nulls
            Assert.True(reader.Read());
            Assert.Equal(DBNull.Value, reader.GetValue(0));
            Assert.Equal(DBNull.Value, reader.GetValue(1));
            Assert.Equal(DBNull.Value, reader.GetValue(2));
            
            // Third row - has values
            Assert.True(reader.Read());
            Assert.Equal("value2", reader.GetValue(0));
            Assert.Equal(2L, (long)reader.GetValue(1));
            Assert.Equal(2.5, (double)reader.GetValue(2));
            
            Assert.False(reader.Read());
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_NullInMiddle_PreservesColumnPositions(string extension)
    {
        // Test for [1, null, 2] - null in middle should not shift values
        var fileName = $"test_null_middle{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(int));
        dt.Columns.Add("Col2", typeof(string));
        dt.Columns.Add("Col3", typeof(int));
        
        dt.Rows.Add(1, DBNull.Value, 2);
        dt.Rows.Add(10, "middle", 20);
        dt.Rows.Add(100, DBNull.Value, 200);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            
            // Row 1: [1, null, 2]
            Assert.True(reader.Read());
            var v0 = reader.GetValue(0);
            var v1 = reader.GetValue(1);
            var v2 = reader.GetValue(2);
            Assert.Equal(1L, (long)reader.GetValue(0));
            Assert.Equal(DBNull.Value, reader.GetValue(1));
            Assert.Equal(2L, (long)reader.GetValue(2));
            
            // Row 2: [10, "middle", 20]
            Assert.True(reader.Read());
            Assert.Equal(10L, (long)reader.GetValue(0));
            Assert.Equal("middle", reader.GetValue(1));
            Assert.Equal(20L, (long)reader.GetValue(2));
            
            // Row 3: [100, null, 200]
            Assert.True(reader.Read());
            Assert.Equal(100L, (long)reader.GetValue(0));
            Assert.Equal(DBNull.Value, reader.GetValue(1));
            Assert.Equal(200L, (long)reader.GetValue(2));
            
            Assert.False(reader.Read());
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_SpecialCharactersInSheetName_Works(string extension)
    {
        var fileName = $"test_sheet_name_special{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            // Sheet names with various characters
            writer.AddSheet("Sheet-Test");
            writer.WriteSheet(dt.CreateDataReader());
            writer.AddSheet("Sheet_Test");
            writer.WriteSheet(dt.CreateDataReader());
            writer.AddSheet("Sheet Test");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            var names = reader.GetScheetNames();
            
            Assert.Equal(3, names.Length);
            Assert.Contains("Sheet-Test", names);
            Assert.Contains("Sheet_Test", names);
            Assert.Contains("Sheet Test", names);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_LargeDataSet_CompletesSuccessfully(string extension)
    {
        var fileName = $"test_large_data{extension}";
        const int rowCount = 10000;
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Id", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Value", typeof(double));
        
        for (int i = 0; i < rowCount; i++)
        {
            dt.Rows.Add(i, $"Name_{i}", i * 0.1);
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
            
            int count = 0;
            while (reader.Read())
            {
                count++;
            }
            // +1 for header row
            Assert.Equal(rowCount + 1, count);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_ManyColumns_CompletesSuccessfully(string extension)
    {
        var fileName = $"test_many_columns{extension}";
        const int columnCount = 100;
        
        DataTable dt = new DataTable();
        for (int i = 0; i < columnCount; i++)
        {
            dt.Columns.Add($"Col{i}", typeof(string));
        }
        
        var rowValues = new object[columnCount];
        for (int i = 0; i < columnCount; i++)
        {
            rowValues[i] = $"Value{i}";
        }
        dt.Rows.Add(rowValues);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Header
            Assert.Equal(columnCount, reader.FieldCount);
            
            Assert.True(reader.Read()); // Data row
            for (int i = 0; i < columnCount; i++)
            {
                Assert.Equal($"Value{i}", reader.GetValue(i));
            }
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_SingleRow_Works(string extension)
    {
        var fileName = $"test_single_row{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("OnlyRow");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Header
            Assert.True(reader.Read()); // Single data row
            Assert.Equal("OnlyRow", reader.GetValue(0));
            Assert.False(reader.Read()); // No more rows
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_SingleColumn_Works(string extension)
    {
        var fileName = $"test_single_column{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("OnlyColumn", typeof(int));
        dt.Rows.Add(1);
        dt.Rows.Add(2);
        dt.Rows.Add(3);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Header
            Assert.Equal(1, reader.FieldCount);
            Assert.Equal("OnlyColumn", reader.GetValue(0));
            
            Assert.True(reader.Read());
            Assert.Equal(1L, reader.GetValue(0)); // GetValue returns long for integers
            Assert.True(reader.Read());
            Assert.Equal(2L, reader.GetValue(0)); // GetValue returns long for integers
            Assert.True(reader.Read());
            Assert.Equal(3L, reader.GetValue(0)); // GetValue returns long for integers
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_UnicodeSheetName_Works(string extension)
    {
        var fileName = $"test_unicode_sheet{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Tablica\u4e16\u754c"); // Chinese characters
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            var names = reader.GetScheetNames();
            
            Assert.Single(names);
            Assert.Contains("Tablica\u4e16\u754c", names);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_MaxInt32_Works(string extension)
    {
        var fileName = $"test_max_int{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Rows.Add(int.MaxValue);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            Assert.True(reader.Read());
            Assert.Equal((long)int.MaxValue, reader.GetValue(0)); // GetValue returns long for integers
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_MinInt32_Works(string extension)
    {
        var fileName = $"test_min_int{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Rows.Add(int.MinValue);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            Assert.True(reader.Read());
            Assert.Equal((long)int.MinValue, reader.GetValue(0)); // GetValue returns long for integers
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_DuplicateColumnNames_Works(string extension)
    {
        var fileName = $"test_duplicate_cols{extension}";
        
        // DataTable doesn't allow duplicate column names, so use a list instead
        List<string> headers = new List<string> { "Col1", "Col1" }; // Duplicate name
        List<TypeCode> typeCodes = new List<TypeCode> { TypeCode.String, TypeCode.String };
        List<object?[]> data = new List<object?[]> { new object?[] { "A", "B" } };

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(headers, typeCodes, data, headers: true);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Header
            Assert.Equal(2, reader.FieldCount);
            
            Assert.True(reader.Read());
            Assert.Equal("A", reader.GetValue(0));
            Assert.Equal("B", reader.GetValue(1));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Write_EmptyDataTable_WithColumns_Works(string extension)
    {
        var fileName = $"test_empty_table{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        // No rows

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            // Should have header only
            Assert.True(reader.Read()); // Header
            Assert.Equal(2, reader.FieldCount);
            Assert.False(reader.Read()); // No data
        }

        File.Delete(fileName);
    }

}
