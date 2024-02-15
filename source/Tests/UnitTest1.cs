using SpreadSheetTasks;
using System.Text;

namespace Tests;

public class UnitTest1
{
    [Fact]
    public void XlsbRead1()
    {
        var path = @"E:\source\repos\SpreadSheetTasks\source\TestFiles\testExcel.xlsb";
        using XlsxOrXlsbReadOrEdit excelFile = ReadFileAndCompare(path);
    }
    
    [Fact]
    public void XlsxRead1()
    {
        var path = @"E:\source\repos\SpreadSheetTasks\source\TestFiles\testExcel.xlsx";
        using XlsxOrXlsbReadOrEdit excelFile = ReadFileAndCompare(path);
    }

    private static XlsxOrXlsbReadOrEdit ReadFileAndCompare(string path)
    {
        XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        var sheetNames = excelFile.GetScheetNames();
        excelFile.ActualSheetName = sheetNames[0];
        object[] row = null;
        StringBuilder sb = new StringBuilder();
        while (excelFile.Read())
        {
            row ??= new object[excelFile.FieldCount];
            excelFile.GetValues(row);
            sb.AppendLine(string.Join('|', row));
        }

        Assert.Equal("""
            A|B||D
            ||ccc|
            |b|ccc|
            |121212||12
            |||
            |||
            |||
            A||1|False
            """,
sb.ToString().Trim());
        return excelFile;
    }

}