using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Tests;

[Collection("Sequential")]
[SupportedOSPlatform("windows")]
public class ExcelComValidation
{
    private static readonly string OutputDir = ResolveTestOutputDir();

    private static string ResolveTestOutputDir()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var testOutput = new DirectoryInfo(Path.Combine(dir.FullName, "test_output"));
            if (testOutput.Exists)
                return testOutput.FullName;
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Cannot find test_output directory.");
    }

    private static bool IsExcelInstalled()
    {
        try { return Type.GetTypeFromProgID("Excel.Application") is not null; }
        catch { return false; }
    }

    private dynamic? _excel;

    private void Start()
    {
        _excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
        _excel.Visible = false;
        _excel.DisplayAlerts = false;
        _excel.ScreenUpdating = false;
    }

    private void Stop()
    {
        if (_excel is not null)
        {
            try { _excel.Quit(); } catch { }
            Marshal.ReleaseComObject(_excel);
            _excel = null;
        }
    }

    private static object? Cell(dynamic ws, int row, int col)
    {
        try { return ws.Cells[row, col].Value; } catch { return null; }
    }

    private void CheckFile(List<string> errors, string fileName, Action<dynamic> assert)
    {
        var path = Path.Combine(OutputDir, fileName);
        if (!File.Exists(path)) { errors.Add($"Missing: {fileName}"); return; }

        dynamic? wb = null;
        try
        {
            wb = _excel.Workbooks.Open(path, false, true);
            assert(wb.Sheets[1]);
            wb.Close(false);
        }
        catch (Exception ex)
        {
            errors.Add($"{fileName}: {ex.Message}");
            if (wb is not null) try { wb.Close(false); } catch { }
        }
    }

    private static bool AutofilterOn(dynamic ws)
    {
        try { return (bool)ws.AutoFilterMode; } catch { }
        try { return ws.AutoFilter is not null; } catch { }
        return false;
    }

    private static void HasFrozenPane(dynamic ws)
    {
        dynamic? w;
        try { w = ws.Windows[1]; } catch { w = ws.Application.ActiveWindow; }
        if (w is not null)
            Assert.Equal(1, (int)w.SplitRow);
    }

    // ──────────────────────────────────────────────────────────────

    [Fact]
    public void Validate_All_Files_Open_In_Excel()
    {
        if (!IsExcelInstalled())
        {
            Console.WriteLine("SKIP — Excel not installed or COM unavailable");
            return;
        }

        var files = Directory.GetFiles(OutputDir, "*.xls?")
            .OrderBy(f => Path.GetFileName(f))
            .ToArray();
        Assert.NotEmpty(files);

        Start();
        try
        {
            int passed = 0, failed = 0;
            var errors = new List<string>();

            foreach (var file in files)
            {
                try
                {
                    dynamic wb = _excel.Workbooks.Open(file, false, true);
                    wb.Close(false);
                    passed++;
                }
                catch (Exception ex)
                {
                    failed++;
                    errors.Add($"{Path.GetFileName(file)}: {ex.Message}");
                }
            }

            Console.WriteLine($"Excel open: {passed}/{passed + failed} OK");
            if (failed > 0)
                Assert.Fail($"Excel open: {passed} passed, {failed} failed\n" + string.Join("\n", errors));
        }
        finally { Stop(); }
    }

    [Fact]
    public void Validate_File_Content_Excel()
    {
        if (!IsExcelInstalled())
        {
            Console.WriteLine("SKIP — Excel not installed or COM unavailable");
            return;
        }

        Start();
        try
        {
            var errors = new List<string>();

            foreach (var ext in new[] { ".xlsx", ".xlsb" })
            {
                CheckFile(errors, $"01_basic_datatable{ext}", ws =>
                {
                    Assert.Equal("ID", Cell(ws, 1, 1)?.ToString());
                    Assert.Equal("Name", Cell(ws, 1, 2)?.ToString());
                    Assert.Equal("Price", Cell(ws, 1, 3)?.ToString());
                    Assert.Equal(1, Convert.ToInt32(Cell(ws, 2, 1)));
                    Assert.Equal("Widget", Cell(ws, 2, 2)?.ToString());
                    Assert.Equal(9.99, Convert.ToDouble(Cell(ws, 2, 3)), 2);
                    Assert.Equal(3, Convert.ToInt32(Cell(ws, 4, 1)));
                    Assert.Equal("Doohickey", Cell(ws, 4, 2)?.ToString());
                    Assert.Equal(0.49, Convert.ToDouble(Cell(ws, 4, 3)), 2);
                    Assert.Null(Cell(ws, 5, 1));
                    Assert.True(AutofilterOn(ws));
                    HasFrozenPane(ws);
                });

                CheckFile(errors, $"03_all_types{ext}", ws =>
                {
                    Assert.Equal("Hello", Cell(ws, 2, 1)?.ToString());
                    Assert.Equal(42, Convert.ToInt32(Cell(ws, 2, 2)));
                    Assert.Equal(3.14, Convert.ToDouble(Cell(ws, 2, 3)), 2);
                    Assert.Equal(true, Convert.ToBoolean(Cell(ws, 2, 4)));
                    Assert.Equal(123.45, Convert.ToDouble(Cell(ws, 2, 6)), 2);
                    Assert.Null(Cell(ws, 4, 1));
                });

                CheckFile(errors, $"05_no_headers{ext}", ws =>
                {
                    Assert.Equal(100, Convert.ToInt32(Cell(ws, 1, 1)));
                    Assert.Equal("data", Cell(ws, 1, 2)?.ToString());
                    Assert.Null(Cell(ws, 2, 1));
                    Assert.True(AutofilterOn(ws));
                });

                CheckFile(errors, $"06_starting_row_col{ext}", ws =>
                {
                    Assert.Equal("A", Cell(ws, 4, 3)?.ToString());
                    Assert.Equal(1, Convert.ToInt32(Cell(ws, 5, 3)));
                    Assert.Equal(2, Convert.ToInt32(Cell(ws, 6, 3)));
                    Assert.Null(Cell(ws, 4, 1));
                    Assert.Null(Cell(ws, 1, 1));
                    Assert.True(AutofilterOn(ws));
                    HasFrozenPane(ws);
                });

                CheckFile(errors, $"09_autofilter{ext}", ws =>
                {
                    Assert.Equal("City", Cell(ws, 1, 1)?.ToString());
                    Assert.Equal("Sales", Cell(ws, 1, 2)?.ToString());
                    Assert.Equal("NYC", Cell(ws, 2, 1)?.ToString());
                    Assert.Equal(100, Convert.ToInt32(Cell(ws, 2, 2)));
                    Assert.Equal("LA", Cell(ws, 3, 1)?.ToString());
                    Assert.Equal(200, Convert.ToInt32(Cell(ws, 3, 2)));
                    Assert.Equal("Chicago", Cell(ws, 4, 1)?.ToString());
                    Assert.Equal(150, Convert.ToInt32(Cell(ws, 4, 2)));
                    Assert.Null(Cell(ws, 5, 1));
                    Assert.True(AutofilterOn(ws));
                });

                CheckFile(errors, $"10_reader_from_list{ext}", ws =>
                {
                    Assert.Equal("ID", Cell(ws, 1, 1)?.ToString());
                    Assert.Equal("Name", Cell(ws, 1, 2)?.ToString());
                    Assert.Equal("Active", Cell(ws, 1, 3)?.ToString());
                    Assert.Equal(1, Convert.ToInt32(Cell(ws, 2, 1)));
                    Assert.Equal("Alice", Cell(ws, 2, 2)?.ToString());
                    Assert.Equal(true, Convert.ToBoolean(Cell(ws, 2, 3)));
                    Assert.Equal(3, Convert.ToInt32(Cell(ws, 4, 1)));
                    Assert.Equal("Charlie", Cell(ws, 4, 2)?.ToString());
                    Assert.True(AutofilterOn(ws));
                });

                CheckFile(errors, $"15_booleans{ext}", ws =>
                {
                    Assert.Equal("Flag", Cell(ws, 1, 1)?.ToString());
                    Assert.Equal("Label", Cell(ws, 1, 2)?.ToString());
                    Assert.Equal(true, Convert.ToBoolean(Cell(ws, 2, 1)));
                    Assert.Equal("Yes", Cell(ws, 2, 2)?.ToString());
                    Assert.Equal(false, Convert.ToBoolean(Cell(ws, 3, 1)));
                    Assert.Equal("No", Cell(ws, 3, 2)?.ToString());
                    Assert.Equal(true, Convert.ToBoolean(Cell(ws, 4, 1)));
                    Assert.Equal("On", Cell(ws, 4, 2)?.ToString());
                    Assert.True(AutofilterOn(ws));
                });

                CheckFile(errors, $"16_null_values{ext}", ws =>
                {
                    Assert.Equal("Col1", Cell(ws, 1, 1)?.ToString());
                    Assert.Equal("Col2", Cell(ws, 1, 2)?.ToString());
                    Assert.Equal("Col3", Cell(ws, 1, 3)?.ToString());
                    Assert.Equal("A", Cell(ws, 2, 1)?.ToString());
                    Assert.Equal(1, Convert.ToInt32(Cell(ws, 2, 2)));
                    Assert.Equal(1.1, Convert.ToDouble(Cell(ws, 2, 3)), 2);
                    Assert.Null(Cell(ws, 3, 1));
                    Assert.Null(Cell(ws, 3, 2));
                    Assert.Null(Cell(ws, 3, 3));
                    Assert.Equal("C", Cell(ws, 4, 1)?.ToString());
                    Assert.Equal(3, Convert.ToInt32(Cell(ws, 4, 2)));
                    Assert.Null(Cell(ws, 4, 3));
                    Assert.True(AutofilterOn(ws));
                });

                CheckFile(errors, $"17_object_array{ext}", ws =>
                {
                    Assert.Equal("ID", Cell(ws, 1, 1)?.ToString());
                    Assert.Equal("Name", Cell(ws, 1, 2)?.ToString());
                    Assert.Equal("Value", Cell(ws, 1, 3)?.ToString());
                    Assert.Equal(1, Convert.ToInt32(Cell(ws, 2, 1)));
                    Assert.Equal("Alice", Cell(ws, 2, 2)?.ToString());
                    Assert.Equal(100.5, Convert.ToDouble(Cell(ws, 2, 3)), 2);
                    Assert.Equal(3, Convert.ToInt32(Cell(ws, 4, 1)));
                    Assert.Equal("Charlie", Cell(ws, 4, 2)?.ToString());
                    Assert.Equal(300.25, Convert.ToDouble(Cell(ws, 4, 3)), 2);
                    Assert.True(AutofilterOn(ws));
                });

                // ── 04_multi_sheet: uses ActiveWorkbook to access both sheets ──
                CheckFile(errors, $"04_multi_sheet{ext}", ws =>
                {
                    dynamic wb = _excel.ActiveWorkbook;
                    Assert.Equal(2, wb.Sheets.Count);
                    dynamic s1 = wb.Sheets[1];
                    Assert.Equal("X", Cell(s1, 1, 1)?.ToString());
                    Assert.Equal(1, Convert.ToInt32(Cell(s1, 2, 1)));
                    Assert.Equal(2, Convert.ToInt32(Cell(s1, 3, 1)));
                    dynamic s2 = wb.Sheets[2];
                    Assert.Equal("Y", Cell(s2, 1, 1)?.ToString());
                    Assert.Equal("aaa", Cell(s2, 2, 1)?.ToString());
                    Assert.Equal("bbb", Cell(s2, 3, 1)?.ToString());
                });
            }

            if (errors.Count > 0)
                Assert.Fail("Content validation failed:\n" + string.Join("\n", errors));

            Console.WriteLine($"Excel content: all {errors.Count} assertion groups passed");
        }
        finally { Stop(); }
    }
}
