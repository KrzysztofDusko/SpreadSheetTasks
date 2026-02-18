using SpreadSheetTasks;
using System.Data;

namespace Tests;

[Collection("Sequential")]
public class ReaderTypeConversionTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetInt32_FromInt32Column_ReturnsValue(string extension)
    {
        var fileName = $"test_int32_value{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Rows.Add(42);

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
            Assert.Equal(42, reader.GetInt32(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetInt32_FromInt64Column_ConvertsAndReturns(string extension)
    {
        var fileName = $"test_int32_from_int64{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("LongCol", typeof(long));
        dt.Rows.Add(100L);

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
            // Int64 should be convertible to Int32
            Assert.Equal(100, reader.GetInt32(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetInt32_FromDoubleColumn_ConvertsAndReturns(string extension)
    {
        var fileName = $"test_int32_from_double{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Rows.Add(42.5);

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
            // Double should be convertible to Int32 (truncated)
            Assert.Equal(42, reader.GetInt32(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetInt64_FromInt64Column_ReturnsValue(string extension)
    {
        var fileName = $"test_int64_value{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("LongCol", typeof(long));
        dt.Rows.Add(9876543210L);

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
            Assert.Equal(9876543210L, reader.GetInt64(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetInt64_FromDoubleColumn_ConvertsAndReturns(string extension)
    {
        var fileName = $"test_int64_from_double{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Rows.Add(123.45);

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
            Assert.Equal(123L, reader.GetInt64(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetDouble_FromDoubleColumn_ReturnsValue(string extension)
    {
        var fileName = $"test_double_value{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Rows.Add(3.14159);

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
            Assert.Equal(3.14159, reader.GetDouble(0), 5);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetDouble_FromIntColumn_ConvertsAndReturns(string extension)
    {
        var fileName = $"test_double_from_int{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Rows.Add(42);

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
            Assert.Equal(42.0, reader.GetDouble(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetDateTime_FromDateTimeColumn_ReturnsValue(string extension)
    {
        var fileName = $"test_datetime_value{extension}";
        
        var testDate = new DateTime(2024, 6, 15, 10, 30, 0);
        
        DataTable dt = new DataTable();
        dt.Columns.Add("DateCol", typeof(DateTime));
        dt.Rows.Add(testDate);

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
            var result = reader.GetDateTime(0);
            Assert.Equal(testDate, result);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetDateTime_FromNonDateTimeColumn_ThrowsInvalidCastException(string extension)
    {
        var fileName = $"test_datetime_invalid{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("StringCol", typeof(string));
        dt.Rows.Add("NotADate");

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
            Assert.Throws<InvalidCastException>(() => reader.GetDateTime(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetInt32_FromStringColumn_ThrowsInvalidCastException(string extension)
    {
        var fileName = $"test_int32_invalid{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("StringCol", typeof(string));
        dt.Rows.Add("NotAnInt");

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
            Assert.Throws<InvalidCastException>(() => reader.GetInt32(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetDouble_FromStringColumn_ThrowsInvalidCastException(string extension)
    {
        var fileName = $"test_double_invalid{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("StringCol", typeof(string));
        dt.Rows.Add("NotADouble");

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
            Assert.Throws<InvalidCastException>(() => reader.GetDouble(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetInt64_FromInt32Column_ReturnsValue(string extension)
    {
        var fileName = $"test_int64_from_int32{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Rows.Add(42);

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
            // Int32 should be readable as Int64
            Assert.Equal(42L, reader.GetInt64(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetDouble_FromInt64Column_ConvertsAndReturns(string extension)
    {
        var fileName = $"test_double_from_int64{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("LongCol", typeof(long));
        dt.Rows.Add(123456789L);

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
            Assert.Equal(123456789.0, reader.GetDouble(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetValues_FillsArrayCorrectly(string extension)
    {
        var fileName = $"test_getvalues{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(int));
        dt.Columns.Add("Col2", typeof(string));
        dt.Columns.Add("Col3", typeof(double));
        dt.Rows.Add(1, "test", 3.14);

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
            
            object[] values = new object[3];
            reader.GetValues(values);
            
            Assert.Equal(1L, values[0]); // GetValue returns long for integers
            Assert.Equal("test", values[1]);
            Assert.Equal(3.14, values[2]);
        }

        File.Delete(fileName);
    }
}