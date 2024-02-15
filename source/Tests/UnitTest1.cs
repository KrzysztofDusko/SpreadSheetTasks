using SpreadSheetTasks;
using System.Text;

namespace Tests;

[Collection("Sequential")]
public class UnitTest1
{
    [Fact]
    public void XlsbRead1()
    {
        var path = @"E:\source\repos\SpreadSheetTasks\source\TestFiles\testExcel.xlsb";
        var res = ReadFileAndCompare(path);

        Assert.Equal("""
            A|B||D
            ||ccc|
            |b|ccc|
            |121212||12
            |||
            |||
            |||
            A||1|False
            """, res);
    }

    [Fact]
    public void XlsxRead1()
    {
        var path = @"E:\source\repos\SpreadSheetTasks\source\TestFiles\testExcel.xlsx";
        var res = ReadFileAndCompare(path);

        Assert.Equal("""
            A|B||D
            ||ccc|
            |b|ccc|
            |121212||12
            |||
            |||
            |||
            A||1|False
            """, res);
    }

    [Fact]
    public void XlsxVsXlsx()
    {
        var pathXlsx = @"E:\source\repos\SpreadSheetTasks\source\TestFiles\testExcel2.xlsx";
        var resXlsx = ReadFileAndCompare(pathXlsx);
        var pathXlsb = @"E:\source\repos\SpreadSheetTasks\source\TestFiles\testExcel2.xlsb";
        var resXlsb = ReadFileAndCompare(pathXlsb);
        Assert.Equal(resXlsx, resXlsb);


        pathXlsx = @"E:\source\repos\SpreadSheetTasks\source\TestFiles\testExcel3.xlsx";
        resXlsx = ReadFileAndCompare(pathXlsx);
        pathXlsb = @"E:\source\repos\SpreadSheetTasks\source\TestFiles\testExcel3.xlsb";
        resXlsb = ReadFileAndCompare(pathXlsb);
        Assert.Equal(resXlsx, resXlsb);
    }


    private static string ReadFileAndCompare(string path)
    {
        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
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


        return sb.ToString().Trim();
    }

}