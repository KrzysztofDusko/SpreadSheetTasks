using System.Data;

namespace Tests;

[Collection("Sequential")]
public class SylvanCompatibilityAdditionalTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Writers_ProduceEquivalentRows_ForMixedScalars(string extension)
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("TextCol", typeof(string));
        dt.Columns.Add("IntCol", typeof(int));
        dt.Columns.Add("DoubleCol", typeof(double));
        dt.Columns.Add("BoolCol", typeof(bool));

        dt.Rows.Add("alpha", 1, 1.5, true);
        dt.Rows.Add("beta", -20, -3.25, false);
        dt.Rows.Add("gamma", DBNull.Value, 0.0, true);

        AssertWriterCompatibility(extension, dt, ["TextCol", "IntCol", "DoubleCol", "BoolCol"]);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Writers_ProduceEquivalentRows_ForUnicodeWhitespaceAndNull(string extension)
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("City", typeof(string));
        dt.Columns.Add("Leading", typeof(string));
        dt.Columns.Add("Trailing", typeof(string));

        dt.Rows.Add("Łódź", "  leading", "trailing  ");
        dt.Rows.Add("zażółć gęślą jaźń", "\tTabbed", "Line1\nLine2");
        dt.Rows.Add(DBNull.Value, "", " ");

        AssertWriterCompatibility(extension, dt, ["City", "Leading", "Trailing"]);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Writers_ProduceEquivalentRows_ForDateTimeAndNullRows(string extension)
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("DateCol", typeof(DateTime));
        dt.Columns.Add("TextCol", typeof(string));
        dt.Columns.Add("IntCol", typeof(int));

        dt.Rows.Add(new DateTime(2024, 1, 1, 12, 30, 0), "first", 10);
        dt.Rows.Add(new DateTime(2024, 6, 15, 0, 0, 0), "second", -5);
        dt.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value);
        dt.Rows.Add(new DateTime(2025, 1, 1, 23, 59, 59), "after", 99);

        AssertWriterCompatibility(extension, dt, ["DateCol", "TextCol", "IntCol"]);
    }

    private static void AssertWriterCompatibility(string extension, DataTable dataTable, string[] headers)
    {
        var spreadPath = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        var sylvanPath = SylvanInteropTestHelpers.CreateTempExcelPath(extension);

        try
        {
            SylvanInteropTestHelpers.WriteWithSpreadSheetTasks(spreadPath, dataTable);
            SylvanInteropTestHelpers.WriteWithSylvan(sylvanPath, dataTable);

            var spreadReadSpread = SylvanInteropTestHelpers.ReadDataRowsWithSpreadSheetTasks(spreadPath);
            var spreadReadSylvan = SylvanInteropTestHelpers.ReadDataRowsWithSpreadSheetTasks(sylvanPath);
            SylvanInteropTestHelpers.AssertRowsEqual(spreadReadSpread, spreadReadSylvan);

            var sylvanReadSpread = SylvanInteropTestHelpers.ReadDataRowsWithSylvan(spreadPath, headers);
            var sylvanReadSylvan = SylvanInteropTestHelpers.ReadDataRowsWithSylvan(sylvanPath, headers);
            SylvanInteropTestHelpers.AssertRowsEqual(sylvanReadSpread, sylvanReadSylvan);
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
}
