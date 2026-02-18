using SpreadSheetTasks;
using Sylvan.Data.Excel;
using System.Data;
using System.Globalization;

namespace Tests;

internal static class SylvanInteropTestHelpers
{
    internal const string SheetName = "Sheet1";

    internal static string CreateTempExcelPath(string extension)
    {
        return Path.Combine(Path.GetTempPath(), $"SpreadSheetTasks_{Guid.NewGuid():N}{extension}");
    }

    internal static void WriteWithSpreadSheetTasks(string path, DataTable dataTable)
    {
        using var writer = ExcelWriter.CreateWriter(path);
        writer.AddSheet(SheetName);
        writer.WriteSheet(dataTable.CreateDataReader());
    }

    internal static void WriteWithSylvan(string path, DataTable dataTable)
    {
        using var writer = ExcelDataWriter.Create(path);
        using var dataReader = dataTable.CreateDataReader();
        writer.Write(dataReader, SheetName);
    }

    internal static List<object?[]> ReadAllRowsWithSpreadSheetTasks(string path)
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(path);
        reader.ActualSheetName = SheetName;

        var rows = new List<object?[]>();
        while (reader.Read())
        {
            object[] row = new object[reader.FieldCount];
            reader.GetValues(row);
            rows.Add(NormalizeRow(row));
        }

        return rows;
    }

    internal static List<object?[]> ReadDataRowsWithSpreadSheetTasks(string path)
    {
        var rows = ReadAllRowsWithSpreadSheetTasks(path);
        if (rows.Count > 0)
        {
            rows.RemoveAt(0); // SpreadSheetTasks reader includes header row.
        }

        return rows;
    }

    internal static List<object?[]> ReadAllRowsWithSylvan(string path)
    {
        using var reader = ExcelDataReader.Create(path);

        var rows = new List<object?[]>();
        while (reader.Read())
        {
            object[] row = new object[reader.FieldCount];
            reader.GetValues(row);
            rows.Add(NormalizeRow(row));
        }

        return rows;
    }

    internal static List<object?[]> ReadDataRowsWithSylvan(string path, IReadOnlyList<string>? headers = null)
    {
        var rows = ReadAllRowsWithSylvan(path);
        if (headers is not null && rows.Count > 0 && IsHeaderRow(rows[0], headers))
        {
            rows.RemoveAt(0);
        }

        return rows;
    }

    internal static void AssertRowsEqual(List<object?[]> expected, List<object?[]> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        for (int i = 0; i < expected.Count; i++)
        {
            Assert.Equal(expected[i], actual[i]);
        }
    }

    internal static bool IsAllNullRow(object?[] row)
    {
        for (int i = 0; i < row.Length; i++)
        {
            if (row[i] is not null)
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsHeaderRow(object?[] row, IReadOnlyList<string> headers)
    {
        if (row.Length < headers.Count)
        {
            return false;
        }

        for (int i = 0; i < headers.Count; i++)
        {
            if (!StringComparer.Ordinal.Equals(row[i] as string, headers[i]))
            {
                return false;
            }
        }

        return true;
    }

    private static object?[] NormalizeRow(object[] row)
    {
        var normalized = new object?[row.Length];
        for (int i = 0; i < row.Length; i++)
        {
            normalized[i] = NormalizeCell(row[i]);
        }

        return normalized;
    }

    private static object? NormalizeCell(object? value)
    {
        if (value is null || value == DBNull.Value)
        {
            return null;
        }

        if (value is sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal)
        {
            return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
        }

        if (value is DateTime dateTime)
        {
            return Math.Round(Convert.ToDecimal(dateTime.ToOADate(), CultureInfo.InvariantCulture), 10);
        }

        if (value is string str && decimal.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedDecimal))
        {
            return parsedDecimal;
        }

        return value;
    }
}
