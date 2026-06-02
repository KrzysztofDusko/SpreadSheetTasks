using SpreadSheetTasks;

string outputDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", ".."));
Console.WriteLine($"Verifying files in: {outputDir}");
Console.WriteLine();

int totalVerified = 0;
int totalErrors = 0;

foreach (var ext in new[] { "xlsx", "xlsb" })
{
    var files = System.IO.Directory.GetFiles(outputDir, $"*.{ext}");
    foreach (var file in files.OrderBy(f => f))
    {
        var fileName = Path.GetFileName(file);
        try
        {
            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(file);

            System.Collections.Generic.IReadOnlyList<string> names = reader.GetSheetNames();
            if (names == null || names.Count == 0)
                throw new Exception("No sheets found");

            int totalRows = 0;
            foreach (var sheet in names)
            {
                reader.ActualSheetName = sheet;
                while (reader.Read())
                    totalRows++;
            }

            totalVerified++;
            var size = new System.IO.FileInfo(file).Length;
            Console.WriteLine($"  [OK] {fileName,-35} sheets={names.Count,2} rows={totalRows,5} size={size,7} B");
        }
        catch (Exception ex)
        {
            totalErrors++;
            Console.WriteLine($"  [FAIL] {fileName,-35} {ex.GetType().Name}: {ex.Message}");
        }
    }
}

Console.WriteLine();
Console.WriteLine(new string('=', 60));
Console.WriteLine($"Verified: {totalVerified} | Errors: {totalErrors}");
Console.WriteLine(new string('=', 60));

if (totalErrors > 0) Environment.Exit(1);
