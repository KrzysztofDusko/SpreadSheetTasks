using System.Data;

namespace Tests;

[Collection("Sequential")]
public class SylvanNullCompatibilityTests
{
    private static readonly string[] _headers = ["Col1", "Col2", "Col3"];

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void SpreadSheetTasksWrite_NullMiddle_IsReadConsistently(string extension)
    {
        var path = SylvanInteropTestHelpers.CreateTempExcelPath(extension);

        try
        {
            var dt = CreateNullFocusedDataTable();
            SylvanInteropTestHelpers.WriteWithSpreadSheetTasks(path, dt);

            var spreadRows = SylvanInteropTestHelpers.ReadDataRowsWithSpreadSheetTasks(path);
            var sylvanRows = SylvanInteropTestHelpers.ReadDataRowsWithSylvan(path, _headers);

            var spreadSnapshot = CaptureNullSnapshot(spreadRows);
            var sylvanSnapshot = CaptureNullSnapshot(sylvanRows);

            Assert.True(spreadSnapshot.HasRow1Null2);
            Assert.True(spreadSnapshot.HasRow100Null200);
            Assert.True(spreadSnapshot.HasRow11After12);

            Assert.Equal(spreadSnapshot.HasRow1Null2, sylvanSnapshot.HasRow1Null2);
            Assert.Equal(spreadSnapshot.HasRow100Null200, sylvanSnapshot.HasRow100Null200);
            Assert.Equal(spreadSnapshot.HasRow11After12, sylvanSnapshot.HasRow11After12);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void SylvanWrite_NullMiddle_IsReadConsistently(string extension)
    {
        var path = SylvanInteropTestHelpers.CreateTempExcelPath(extension);

        try
        {
            var dt = CreateNullFocusedDataTable();
            SylvanInteropTestHelpers.WriteWithSylvan(path, dt);

            var spreadRows = SylvanInteropTestHelpers.ReadDataRowsWithSpreadSheetTasks(path);
            var sylvanRows = SylvanInteropTestHelpers.ReadDataRowsWithSylvan(path, _headers);

            var spreadSnapshot = CaptureNullSnapshot(spreadRows);
            var sylvanSnapshot = CaptureNullSnapshot(sylvanRows);

            Assert.True(sylvanSnapshot.HasRow1Null2);
            Assert.True(sylvanSnapshot.HasRow100Null200);
            Assert.True(sylvanSnapshot.HasRow11After12);

            Assert.Equal(sylvanSnapshot.HasRow1Null2, spreadSnapshot.HasRow1Null2);
            Assert.Equal(sylvanSnapshot.HasRow100Null200, spreadSnapshot.HasRow100Null200);
            Assert.Equal(sylvanSnapshot.HasRow11After12, spreadSnapshot.HasRow11After12);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void SpreadReader_SpreadAndSylvanWrite_ShouldProduceSameDataRows(string extension)
    {
        var spreadPath = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        var sylvanPath = SylvanInteropTestHelpers.CreateTempExcelPath(extension);

        try
        {
            var dt = CreateNullFocusedDataTable();
            SylvanInteropTestHelpers.WriteWithSpreadSheetTasks(spreadPath, dt);
            SylvanInteropTestHelpers.WriteWithSylvan(sylvanPath, dt);

            var spreadWriteRows = SylvanInteropTestHelpers.ReadDataRowsWithSpreadSheetTasks(spreadPath);
            var sylvanWriteRows = SylvanInteropTestHelpers.ReadDataRowsWithSpreadSheetTasks(sylvanPath);

            var spreadSnapshot = CaptureNullSnapshot(spreadWriteRows);
            var sylvanSnapshot = CaptureNullSnapshot(sylvanWriteRows);

            Assert.Equal(spreadSnapshot, sylvanSnapshot);
        }
        finally
        {
            if (File.Exists(spreadPath))
            {
                File.Delete(spreadPath);
            }

            if (File.Exists(sylvanPath))
            {
                File.Delete(sylvanPath);
            }
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void SylvanReader_SpreadAndSylvanWrite_ShouldProduceSameDataRows(string extension)
    {
        var spreadPath = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        var sylvanPath = SylvanInteropTestHelpers.CreateTempExcelPath(extension);

        try
        {
            var dt = CreateNullFocusedDataTable();
            SylvanInteropTestHelpers.WriteWithSpreadSheetTasks(spreadPath, dt);
            SylvanInteropTestHelpers.WriteWithSylvan(sylvanPath, dt);

            var spreadWriteRows = SylvanInteropTestHelpers.ReadDataRowsWithSylvan(spreadPath, _headers);
            var sylvanWriteRows = SylvanInteropTestHelpers.ReadDataRowsWithSylvan(sylvanPath, _headers);

            var spreadSnapshot = CaptureNullSnapshot(spreadWriteRows);
            var sylvanSnapshot = CaptureNullSnapshot(sylvanWriteRows);

            Assert.Equal(spreadSnapshot, sylvanSnapshot);
        }
        finally
        {
            if (File.Exists(spreadPath))
            {
                File.Delete(spreadPath);
            }

            if (File.Exists(sylvanPath))
            {
                File.Delete(sylvanPath);
            }
        }
    }

    private static DataTable CreateNullFocusedDataTable()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(int));
        dt.Columns.Add("Col2", typeof(string));
        dt.Columns.Add("Col3", typeof(int));

        dt.Rows.Add(1, DBNull.Value, 2);                    // [1, null, 2]
        dt.Rows.Add(10, "middle", 20);                      // non-null row
        dt.Rows.Add(100, DBNull.Value, 200);                // second null-in-middle
        dt.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value); // all-null row
        dt.Rows.Add(11, "after_nulls", 12);                 // detects value leakage

        return dt;
    }

    private static NullScenarioSnapshot CaptureNullSnapshot(List<object?[]> rows)
    {
        return new NullScenarioSnapshot(
            HasRow(rows, 1m, null, 2m),
            HasRow(rows, 100m, null, 200m),
            HasRow(rows, 11m, "after_nulls", 12m),
            rows.Any(SylvanInteropTestHelpers.IsAllNullRow));
    }

    private static bool HasRow(List<object?[]> rows, object? c1, object? c2, object? c3)
    {
        return rows.Any(r =>
            r.Length >= 3 &&
            Equals(r[0], c1) &&
            Equals(r[1], c2) &&
            Equals(r[2], c3));
    }

    private sealed record NullScenarioSnapshot(
        bool HasRow1Null2,
        bool HasRow100Null200,
        bool HasRow11After12,
        bool HasAllNullRow);
}
