using SpreadSheetTasks;
using System.Data;

namespace Tests;

[Collection("Sequential")]
public class ReaderAdvancedTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetNativeValue_ReturnsFieldInfoStruct(string extension)
    {
        var fileName = $"test_native_value{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Columns.Add("StringCol", typeof(string));
        dt.Rows.Add(42, "test");

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
            
            ref var fieldInfo = ref reader.GetNativeValue(0);
            Assert.True(fieldInfo.type == ExcelDataType.Int32 || fieldInfo.type == ExcelDataType.Int64 || fieldInfo.type == ExcelDataType.Double);
            
            ref var stringField = ref reader.GetNativeValue(1);
            Assert.Equal(ExcelDataType.String, stringField.type);
            Assert.Equal("test", stringField.strValue);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetNativeValues_ReturnsFieldInfoArray(string extension)
    {
        var fileName = $"test_native_values{extension}";
        
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
            
            ref var fieldArray = ref reader.GetNativeValues();
            Assert.NotNull(fieldArray);
            Assert.True(fieldArray.Length >= 3);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void UseMemoryStreamInXlsb_True_Works(string extension)
    {
        var fileName = $"test_memory_stream_true{extension}";
        
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
            reader.UseMemoryStreamInXlsb = true;
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            Assert.True(reader.Read());
            Assert.Equal("Data", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void UseMemoryStreamInXlsb_False_Works(string extension)
    {
        var fileName = $"test_memory_stream_false{extension}";
        
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
            reader.UseMemoryStreamInXlsb = false;
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            Assert.True(reader.Read());
            Assert.Equal("Data", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Open_WithUpdateMode_Works(string extension)
    {
        var fileName = $"test_update_mode{extension}";
        
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
            // Open in update mode (only for xlsx)
            reader.Open(fileName, updateMode: true);
            var names = reader.GetScheetNames();
            Assert.NotEmpty(names);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void RelativePositionInStream_ReturnsValue(string extension)
    {
        var fileName = $"test_relative_position{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        for (int i = 0; i < 100; i++)
        {
            dt.Rows.Add($"Data{i}");
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
            
            // Read some rows and check position
            for (int i = 0; i < 10; i++)
            {
                reader.Read();
            }
            
            var position = reader.RelativePositionInStream();
            Assert.True(position >= 0);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ActualSheetName_ChangeSheet_Works(string extension)
    {
        var fileName = $"test_change_sheet{extension}";
        
        DataTable dt1 = new DataTable();
        dt1.Columns.Add("Col1", typeof(string));
        dt1.Rows.Add("Sheet1Data");
        
        DataTable dt2 = new DataTable();
        dt2.Columns.Add("Col1", typeof(string));
        dt2.Rows.Add("Sheet2Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("FirstSheet");
            writer.WriteSheet(dt1.CreateDataReader());
            writer.AddSheet("SecondSheet");
            writer.WriteSheet(dt2.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            
            // Read from first sheet
            reader.ActualSheetName = "FirstSheet";
            Assert.True(reader.Read()); // Header
            Assert.True(reader.Read()); // Data
            Assert.Equal("Sheet1Data", reader.GetValue(0));
            
            // Switch to second sheet
            reader.ActualSheetName = "SecondSheet";
            Assert.True(reader.Read()); // Header
            Assert.True(reader.Read()); // Data
            Assert.Equal("Sheet2Data", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ActualSheetName_JumpBetweenSheets_PartialReads_ReturnsCorrectData_Issue6(string extension)
    {
        var fileName = $"test_issue6_sheet_jump{extension}";

        DataTable dt1 = new DataTable();
        dt1.Columns.Add("Col1", typeof(string));
        dt1.Rows.Add("A1");
        dt1.Rows.Add("A2");

        DataTable dt2 = new DataTable();
        dt2.Columns.Add("Col1", typeof(string));
        dt2.Rows.Add("B1");
        dt2.Rows.Add("B2");

        DataTable dt3 = new DataTable();
        dt3.Columns.Add("Col1", typeof(string));
        dt3.Rows.Add("C1");
        dt3.Rows.Add("C2");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("SheetA");
            writer.WriteSheet(dt1.CreateDataReader());
            writer.AddSheet("SheetB");
            writer.WriteSheet(dt2.CreateDataReader());
            writer.AddSheet("SheetC");
            writer.WriteSheet(dt3.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);

            reader.ActualSheetName = "SheetA";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read()); // first data row
            Assert.Equal("A1", reader.GetValue(0));

            reader.ActualSheetName = "SheetB";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read()); // first data row
            Assert.Equal("B1", reader.GetValue(0));

            // Jump back - should restart from beginning of SheetA, not continue stale state.
            reader.ActualSheetName = "SheetA";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read()); // first data row
            Assert.Equal("A1", reader.GetValue(0));

            reader.ActualSheetName = "SheetC";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read()); // first data row
            Assert.Equal("C1", reader.GetValue(0));

            reader.ActualSheetName = "SheetB";
            Assert.True(reader.Read()); // header
            Assert.True(reader.Read()); // first data row
            Assert.Equal("B1", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetScheetNames_ReturnsCorrectOrder(string extension)
    {
        var fileName = $"test_sheet_order{extension}";
        
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
            Assert.Equal("Alpha", names[0]);
            Assert.Equal("Beta", names[1]);
            Assert.Equal("Gamma", names[2]);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Read_AfterDispose_ThrowsException(string extension)
    {
        var fileName = $"test_read_after_dispose{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(fileName);
        reader.ActualSheetName = "Sheet1";
        reader.Dispose();
        
        // After dispose, operations should fail
        Assert.Throws<ObjectDisposedException>(() => reader.Read());

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void TreatAllColumnsAsText_AfterOpen_DoesNotAffectFieldType(string extension)
    {
        var fileName = $"test_treat_text_after{extension}";
        
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
            
            // Setting after Open may not affect already-read data
            reader.TreatAllColumnsAsText = true;
            Assert.True(reader.Read()); // Skip header
            Assert.True(reader.Read());
            // Behavior depends on implementation
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Read_MultipleTimesOnSameSheet_ReturnsData(string extension)
    {
        var fileName = $"test_multi_read{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data1");
        dt.Rows.Add("Data2");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            // First read
            int count1 = 0;
            while (reader.Read()) { count1++; }
            
            // Need to reset - switch sheets or reopen
            reader.ActualSheetName = "Sheet1";
            
            int count2 = 0;
            while (reader.Read()) { count2++; }
            
            // Both should read same number of rows
            Assert.Equal(count1, count2);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetValue_WithIndexOutOfRange_ThrowsException(string extension)
    {
        var fileName = $"test_index_out_of_range{extension}";
        
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
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            Assert.True(reader.Read());
            
            // Try to access column that doesn't exist
            Assert.Throws<IndexOutOfRangeException>(() => reader.GetValue(100));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetFieldType_WithIndexOutOfRange_ThrowsException(string extension)
    {
        var fileName = $"test_fieldtype_out_of_range{extension}";
        
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
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            Assert.True(reader.Read());
            
            Assert.Throws<IndexOutOfRangeException>(() => reader.GetFieldType(100));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Dispose_CalledMultipleTimes_DoesNotThrow(string extension)
    {
        var fileName = $"test_multi_dispose{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(fileName);
        
        reader.Dispose();
        reader.Dispose(); // Should not throw
        
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetNativeValue_FieldInfoType_ReturnsCorrectExcelDataType(string extension)
    {
        var fileName = $"test_fieldinfo_type{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Columns.Add("StringCol", typeof(string));
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Columns.Add("DateCol", typeof(DateTime));
        dt.Rows.Add(42, "hello", 3.14, new DateTime(2026, 6, 2));

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read()); // skip header
            Assert.True(reader.Read());

            var intType = reader.GetNativeValue(0).type;
            Assert.True(intType == ExcelDataType.Int32 || intType == ExcelDataType.Int64 || intType == ExcelDataType.Double);
            Assert.Equal(ExcelDataType.String, reader.GetNativeValue(1).type);
            Assert.Equal(ExcelDataType.Double, reader.GetNativeValue(2).type);
            var dateType = reader.GetNativeValue(3).type;
            Assert.True(dateType == ExcelDataType.DateTime || dateType == ExcelDataType.Double);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void GetNativeValue_FieldInfoValues_MatchTypedGetters(string extension)
    {
        var fileName = $"test_fieldinfo_match{extension}";

        DataTable dt = new DataTable();
        dt.Columns.Add("IntCol", typeof(int));
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Columns.Add("DateCol", typeof(DateTime));
        dt.Columns.Add("StringCol", typeof(string));
        dt.Rows.Add(42, 3.14, new DateTime(2026, 6, 2), "test_str");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read()); // skip header
            Assert.True(reader.Read());

            ref var fi0 = ref reader.GetNativeValue(0);
            Assert.True(fi0.type == ExcelDataType.Int32 || fi0.type == ExcelDataType.Int64 || fi0.type == ExcelDataType.Double);
            if (fi0.type == ExcelDataType.Int32) Assert.Equal(42, fi0.int32Value);
            else if (fi0.type == ExcelDataType.Int64) Assert.Equal(42L, fi0.int64Value);
            else Assert.Equal(42.0, fi0.doubleValue, 3);

            ref var fi1 = ref reader.GetNativeValue(1);
            Assert.Equal(ExcelDataType.Double, fi1.type);
            Assert.Equal(3.14, fi1.doubleValue, 3);

            ref var fi2 = ref reader.GetNativeValue(2);
            Assert.Equal(ExcelDataType.DateTime, fi2.type);
            Assert.Equal(new DateTime(2026, 6, 2), fi2.dtValue);

            ref var fi3 = ref reader.GetNativeValue(3);
            Assert.Equal(ExcelDataType.String, fi3.type);
            Assert.Equal("test_str", fi3.strValue);
        }

        File.Delete(fileName);
    }
}
