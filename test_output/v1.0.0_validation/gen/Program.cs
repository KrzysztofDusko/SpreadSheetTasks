using System.Data;
using SpreadSheetTasks;

string outputDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", ".."));
Console.WriteLine($"Generating test files to: {outputDir}");
Console.WriteLine();

int totalFiles = 0;
int totalErrors = 0;

void Generate(string name, Action<string> build, bool includeXlsb = true)
{
    foreach (var ext in includeXlsb ? new[] { ".xlsx", ".xlsb" } : new[] { ".xlsx" })
    {
        var path = Path.Combine(outputDir, $"{name}{ext}");
        try
        {
            build(path);
            totalFiles++;
            Console.WriteLine($"  [OK] {Path.GetFileName(path)}");
        }
        catch (Exception ex)
        {
            totalErrors++;
            Console.WriteLine($"  [FAIL] {Path.GetFileName(path)}: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine($"         {ex.StackTrace?.Split('\n')[0]?.Trim()}");
        }
    }
}

// ==========================================
// 01 - Basic data: strings + numbers
// ==========================================
Console.WriteLine("01_basic_data");
Generate("01_basic_data", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("Name", typeof(string));
    dt.Columns.Add("Age", typeof(int));
    dt.Columns.Add("Salary", typeof(double));
    dt.Columns.Add("Active", typeof(bool));

    dt.Rows.Add("Alice", 30, 75000.50, true);
    dt.Rows.Add("Bob", 25, 62000.00, false);
    dt.Rows.Add("Charlie", 35, 88000.75, true);
    dt.Rows.Add("Diana", 28, 71000.25, true);
    dt.Rows.Add("Eve", 32, 95000.00, false);

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("Employees");
    writer.WriteSheet(dt);
});

// ==========================================
// 02 - Multiple sheets
// ==========================================
Console.WriteLine("02_multiple_sheets");
Generate("02_multiple_sheets", path =>
{
    using var writer = ExcelWriter.CreateWriter(path);

    var dt1 = new DataTable();
    dt1.Columns.Add("Product", typeof(string));
    dt1.Columns.Add("Price", typeof(double));
    dt1.Rows.Add("Widget", 9.99);
    dt1.Rows.Add("Gadget", 19.99);
    dt1.Rows.Add("Doohickey", 4.99);

    var dt2 = new DataTable();
    dt2.Columns.Add("Region", typeof(string));
    dt2.Columns.Add("Sales", typeof(int));
    dt2.Rows.Add("North", 1000);
    dt2.Rows.Add("South", 1500);
    dt2.Rows.Add("East", 800);
    dt2.Rows.Add("West", 1200);

    var dt3 = new DataTable();
    dt3.Columns.Add("Month", typeof(string));
    dt3.Columns.Add("Revenue", typeof(double));
    dt3.Rows.Add("January", 50000.00);
    dt3.Rows.Add("February", 52000.00);

    writer.AddSheet("Products");
    writer.WriteSheet(dt1);
    writer.AddSheet("Regions");
    writer.WriteSheet(dt2);
    writer.AddSheet("Monthly");
    writer.WriteSheet(dt3);
});

// ==========================================
// 03 - Dates and booleans
// ==========================================
Console.WriteLine("03_dates_and_booleans");
Generate("03_dates_and_booleans", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("Event", typeof(string));
    dt.Columns.Add("Date", typeof(DateTime));
    dt.Columns.Add("Time", typeof(DateTime));
    dt.Columns.Add("Confirmed", typeof(bool));
    dt.Columns.Add("Cancelled", typeof(bool));

    dt.Rows.Add("Meeting", new DateTime(2025, 6, 15), new DateTime(2025, 6, 15, 9, 30, 0), true, false);
    dt.Rows.Add("Workshop", new DateTime(2025, 7, 1), new DateTime(2025, 7, 1, 14, 0, 0), true, false);
    dt.Rows.Add("Conference", new DateTime(2025, 9, 10), new DateTime(2025, 9, 10, 8, 0, 0), false, true);
    dt.Rows.Add("Webinar", new DateTime(2025, 8, 20), new DateTime(2025, 8, 20, 11, 0, 0), true, false);
    dt.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value);

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("Events");
    writer.WriteSheet(dt);
});

// ==========================================
// 04 - Large dataset (1000 rows)
// ==========================================
Console.WriteLine("04_large_dataset");
Generate("04_large_dataset", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("ID", typeof(int));
    dt.Columns.Add("Value", typeof(double));
    dt.Columns.Add("Category", typeof(string));
    dt.Columns.Add("Flag", typeof(bool));

    var rng = new Random(42);
    for (int i = 0; i < 1000; i++)
    {
        dt.Rows.Add(i + 1, Math.Round(rng.NextDouble() * 1000, 2), $"Cat{(i % 5) + 1}", i % 3 == 0);
    }

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("LargeData");
    writer.WriteSheet(dt);
});

// ==========================================
// 05 - AutoFilter
// ==========================================
Console.WriteLine("05_autofilter");
Generate("05_autofilter", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("Country", typeof(string));
    dt.Columns.Add("City", typeof(string));
    dt.Columns.Add("Population", typeof(int));
    dt.Columns.Add("Area", typeof(double));

    dt.Rows.Add("USA", "New York", 8_400_000, 783.8);
    dt.Rows.Add("USA", "Los Angeles", 3_900_000, 1302.0);
    dt.Rows.Add("UK", "London", 8_900_000, 1572.0);
    dt.Rows.Add("Japan", "Tokyo", 13_900_000, 2194.0);
    dt.Rows.Add("France", "Paris", 2_100_000, 105.4);
    dt.Rows.Add("Germany", "Berlin", 3_600_000, 891.8);
    dt.Rows.Add("Brazil", "Sao Paulo", 12_300_000, 1521.0);
    dt.Rows.Add("India", "Mumbai", 12_400_000, 603.4);

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("Cities");
    writer.WriteSheet(dt, doAutofilter: true);
});

// ==========================================
// 06 - Custom DocPropertyProgramName
// ==========================================
Console.WriteLine("06_custom_properties");
Generate("06_custom_properties", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("Item", typeof(string));
    dt.Columns.Add("Qty", typeof(int));
    dt.Rows.Add("Alpha", 10);
    dt.Rows.Add("Beta", 20);

    using var writer = ExcelWriter.CreateWriter(path);
    writer.DocPropertyProgramName = "MyCustomApp v1.0.0";
    writer.AddSheet("Data");
    writer.WriteSheet(dt);
});

// ==========================================
// 07 - No headers
// ==========================================
Console.WriteLine("07_no_headers");
Generate("07_no_headers", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("C1", typeof(string));
    dt.Columns.Add("C2", typeof(int));
    dt.Columns.Add("C3", typeof(double));
    dt.Rows.Add("Row1", 100, 1.5);
    dt.Rows.Add("Row2", 200, 2.5);
    dt.Rows.Add("Row3", 300, 3.5);

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("NoHeaders");
    writer.WriteSheet(dt, headers: false);
});

// ==========================================
// 08 - Offset starting row/column
// ==========================================
Console.WriteLine("08_offset_start");
Generate("08_offset_start", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("A", typeof(string));
    dt.Columns.Add("B", typeof(int));
    dt.Rows.Add("X", 1);
    dt.Rows.Add("Y", 2);
    dt.Rows.Add("Z", 3);

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("Offset");
    writer.WriteSheet(dt, startingRow: 3, startingColumn: 2);
});

// ==========================================
// 09 - Null and empty values
// ==========================================
Console.WriteLine("09_null_and_empty");
Generate("09_null_and_empty", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("StringCol", typeof(string));
    dt.Columns.Add("IntCol", typeof(int));
    dt.Columns.Add("DoubleCol", typeof(double));
    dt.Columns.Add("BoolCol", typeof(bool));
    dt.Columns.Add("DateCol", typeof(DateTime));

    dt.Rows.Add("Hello", 42, 3.14, true, new DateTime(2025, 1, 1));
    dt.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value);
    dt.Rows.Add("", 0, 0.0, false, DBNull.Value);
    dt.Rows.Add("World", DBNull.Value, DBNull.Value, DBNull.Value, new DateTime(2025, 12, 31));
    dt.Rows.Add(DBNull.Value, 99, 2.71, true, DBNull.Value);

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("Nulls");
    writer.WriteSheet(dt);
});

// ==========================================
// 10 - Mixed types
// ==========================================
Console.WriteLine("10_mixed_types");
Generate("10_mixed_types", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("String", typeof(string));
    dt.Columns.Add("Int16", typeof(short));
    dt.Columns.Add("Int32", typeof(int));
    dt.Columns.Add("Int64", typeof(long));
    dt.Columns.Add("Float", typeof(float));
    dt.Columns.Add("Double", typeof(double));
    dt.Columns.Add("Decimal", typeof(decimal));
    dt.Columns.Add("Boolean", typeof(bool));
    dt.Columns.Add("DateTime", typeof(DateTime));
    dt.Columns.Add("Guid", typeof(Guid));

    dt.Rows.Add(
        "Test", (short)100, 200, 300L, 1.5f, 2.5, 3.5m,
        true, new DateTime(2025, 6, 1, 12, 0, 0),
        Guid.NewGuid()
    );
    dt.Rows.Add(
        "Row2", (short)400, 500, 600L, 7.5f, 8.5, 9.5m,
        false, new DateTime(2025, 6, 2, 15, 30, 0),
        Guid.NewGuid()
    );

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("Mixed");
    writer.WriteSheet(dt);
});

// ==========================================
// 11 - WriteSheet(object?[][]) - basic (auto-generated headers)
// ==========================================
Console.WriteLine("11_object_array");
Generate("11_object_array", path =>
{
    object?[][] data =
    [
        ["Header1", "Header2", "Header3"],
        [42, "Text value", 3.14],
        [99, "Another text", 2.71],
        [123, "More data", 1.41],
        [-5, "Negative row", 0.0],
        [0, "Zero row", 100.0],
    ];

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("ObjectArray");
    writer.WriteSheet(data);
});

// ==========================================
// 11b - WriteSheet(object?[][], string[] headers) - new API with custom headers
// ==========================================
Console.WriteLine("11b_object_array_headers");
Generate("11b_object_array_headers", path =>
{
    object?[][] data =
    [
        [42, "Alice", 75000.50],
        [99, "Bob", 62000.00],
        [123, "Charlie", 88000.75],
    ];

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("WithHeaders");
    writer.WriteSheet(data, new[] { "ID", "Name", "Salary" });
});

// ==========================================
// 12 - WriteSheet(FormattedCell?[][]) - new API
// ==========================================
Console.WriteLine("12_formatted_cells");
Generate("12_formatted_cells", path =>
{
    FormattedCell?[][] data = new FormattedCell?[6][];
    data[0] =
    [
        new FormattedCell("Thousands", F.THOUSANDS_SEP),
        new FormattedCell("Currency", F.CURRENCY_PLN),
        new FormattedCell("Percent", F.PERCENTAGE),
        new FormattedCell("Text", F.TEXT),
    ];
    data[1] =
    [
        new FormattedCell(1234567, F.THOUSANDS_SEP),
        new FormattedCell(99.99, F.CURRENCY_PLN),
        new FormattedCell(0.25, F.PERCENTAGE),
        new FormattedCell("00123", F.LEADING_ZEROS),
    ];
    data[2] =
    [
        new FormattedCell("DateShort", F.DATE_SHORT),
        new FormattedCell("DateLong", F.DATE_LONG),
        new FormattedCell("DateISO", F.DATE_ISO),
        new FormattedCell("DateTime", F.DATETIME_24H),
    ];
    data[3] =
    [
        new FormattedCell(new DateTime(2025, 6, 15), F.DATE_SHORT),
        new FormattedCell(new DateTime(2025, 7, 4), F.DATE_LONG),
        new FormattedCell(new DateTime(2025, 8, 20), F.DATE_ISO),
        new FormattedCell(new DateTime(2025, 9, 10, 14, 30, 0), F.DATETIME_24H),
    ];
    data[4] =
    [
        new FormattedCell("TwoDecimals", F.TWO_DECIMALS),
        new FormattedCell("Scientific", F.SCIENTIFIC),
        new FormattedCell("Time", F.TIME_HH_MM_SS),
        new FormattedCell("TimeMs", F.TIME_MS),
    ];
    data[5] =
    [
        new FormattedCell(1234.5678, F.TWO_DECIMALS),
        new FormattedCell(0.00123, F.SCIENTIFIC),
        new FormattedCell(new DateTime(2025, 1, 1, 9, 5, 30), F.TIME_HH_MM_SS),
        new FormattedCell(new DateTime(2025, 1, 1, 23, 59, 59, 123), F.TIME_MS),
    ];

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("Formatted");
    writer.WriteSheet(data);
});

// ==========================================
// 12b - WriteSheet(FormattedCell?[][], string[] headers) - new API with custom headers
// ==========================================
Console.WriteLine("12b_formatted_cells_headers");
Generate("12b_formatted_cells_headers", path =>
{
    FormattedCell?[][] data =
    [
        [new FormattedCell(1234567, F.THOUSANDS_SEP), new FormattedCell(99.99, F.CURRENCY_PLN), new FormattedCell(0.25, F.PERCENTAGE)],
        [new FormattedCell(new DateTime(2025, 6, 15), F.DATE_SHORT), new FormattedCell(new DateTime(2025, 9, 10, 14, 30, 0), F.DATETIME_24H), new FormattedCell(0.00123, F.SCIENTIFIC)],
        [new FormattedCell(42, F.THOUSANDS_SEP), new FormattedCell(123.45, F.CURRENCY_PLN), new FormattedCell(0.75, F.PERCENTAGE)],
    ];

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("FormattedHeaders");
    writer.WriteSheet(data, new[] { "Amount", "Price PLN", "Ratio %" });
});

// ==========================================
// 13 - SuppressYear1000Dates
// ==========================================
Console.WriteLine("13_suppress_year1000");
Generate("13_suppress_year1000", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("Description", typeof(string));
    dt.Columns.Add("DateValue", typeof(DateTime));
    dt.Columns.Add("EmptyDate", typeof(DateTime));

    dt.Rows.Add("Normal date", new DateTime(2025, 3, 15), DBNull.Value);
    dt.Rows.Add("Old date", new DateTime(1999, 12, 31), DBNull.Value);
    dt.Rows.Add("Future date", new DateTime(2030, 1, 1), DBNull.Value);
    dt.Rows.Add("Only DBNull", DBNull.Value, DBNull.Value);

    using var writer = ExcelWriter.CreateWriter(path);
    writer.SuppressYear1000Dates = true;
    writer.AddSheet("Dates");
    writer.WriteSheet(dt);
});

// ==========================================
// 14 - Single row / edge case
// ==========================================
Console.WriteLine("14_edge_cases");
Generate("14_edge_cases", path =>
{
    using var writer = ExcelWriter.CreateWriter(path);

    // Single row, no headers
    var dt1 = new DataTable();
    dt1.Columns.Add("X", typeof(string));
    dt1.Rows.Add("OnlyOne");
    writer.AddSheet("SingleRow");
    writer.WriteSheet(dt1, headers: false);

    // Empty DataTable (only headers)
    var dt2 = new DataTable();
    dt2.Columns.Add("ColA", typeof(string));
    dt2.Columns.Add("ColB", typeof(int));
    writer.AddSheet("HeadersOnly");
    writer.WriteSheet(dt2);

    // Very long string
    var dt3 = new DataTable();
    dt3.Columns.Add("LongText", typeof(string));
    dt3.Rows.Add(new string('A', 1000));
    writer.AddSheet("LongString");
    writer.WriteSheet(dt3);
});

// ==========================================
// 15 - Polish characters / Unicode
// ==========================================
Console.WriteLine("15_unicode");
Generate("15_unicode", path =>
{
    var dt = new DataTable();
    dt.Columns.Add("Polish", typeof(string));
    dt.Columns.Add("Description", typeof(string));
    dt.Rows.Add("Zażółć gęślą jaźń", "Polish pangram");
    dt.Rows.Add("W Szczebrzeszynie chrząszcz brzmi w trzcinie", "Famous tongue twister");
    dt.Rows.Add("Pchnąć w tę łódź jeża lub ośm skrzyń fig", "Another pangram");
    dt.Rows.Add("ストリーム", "Japanese");
    dt.Rows.Add("StreamWriter", "English");
    dt.Rows.Add("🍕🎉💻", "Emoji");

    using var writer = ExcelWriter.CreateWriter(path);
    writer.AddSheet("Unicode");
    writer.WriteSheet(dt);
});

// ==========================================
// Summary
// ==========================================
Console.WriteLine();
Console.WriteLine(new string('=', 50));
Console.WriteLine($"Total files generated: {totalFiles}");
Console.WriteLine($"Total errors: {totalErrors}");
Console.WriteLine($"Output directory: {outputDir}");
Console.WriteLine(new string('=', 50));

if (totalErrors > 0)
{
    Environment.Exit(1);
}
