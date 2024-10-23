using SpreadSheetTasks;
using System.Data;
using System.Security.Cryptography;
using System.Text;

namespace Tests;

[Collection("Sequential")]
public class UnitTest1
{
    [Fact]
    public void XlsbRead1()
    {
        var path = @"D:\DEV\source\repos\SpreadSheetTasks\source\TestFiles\testExcel.xlsb";
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
        var path = @"D:\DEV\source\repos\SpreadSheetTasks\source\TestFiles\testExcel.xlsx";
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
        var pathXlsx = @"D:\DEV\source\repos\SpreadSheetTasks\source\TestFiles\testExcel2.xlsx";
        var resXlsx = ReadFileAndCompare(pathXlsx);
        var pathXlsb = @"D:\DEV\source\repos\SpreadSheetTasks\source\TestFiles\testExcel2.xlsb";
        var resXlsb = ReadFileAndCompare(pathXlsb);
        Assert.Equal(resXlsx, resXlsb);


        pathXlsx = @"D:\DEV\source\repos\SpreadSheetTasks\source\TestFiles\testExcel3.xlsx";
        resXlsx = ReadFileAndCompare(pathXlsx);
        pathXlsb = @"D:\DEV\source\repos\SpreadSheetTasks\source\TestFiles\testExcel3.xlsb";
        resXlsb = ReadFileAndCompare(pathXlsb);
        Assert.Equal(resXlsx, resXlsb);
    }


    [Fact]
    public void XlsxWriteT1()
    {
        var dt = GetDataTable();
        {
            using var excel = new XlsxWriter("file1.xlsx");
            excel.AddSheet("sheetName");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX1");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX2");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
        }
        {
            using var fs = File.Open("file2.xlsx", System.IO.FileMode.Create);
            var excel = new XlsxWriter(fs);
            excel.AddSheet("sheetName");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX1");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX2");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.Dispose();
        }
        {
            using var memoryStream = new MemoryStream();
            var excel = new XlsxWriter(memoryStream);
            excel.AddSheet("sheetName");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX1");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX2");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.Dispose();

            using var fileStream = File.Open("file3.xlsx", FileMode.Create);
            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.CopyTo(fileStream);
        }


        var str1 = ReadFileAndCompare("file1.xlsx");
        var str2 = ReadFileAndCompare("file2.xlsx");
        var str3 = ReadFileAndCompare("file3.xlsx");
        Assert.Equal(str1, str2);
        Assert.Equal(str1, str3);
        Assert.Equal(str2, str3);


        var sha1 = SHA256.HashData(new FileStream("file1.xlsx", FileMode.Open));
        var sha2 = SHA256.HashData(new FileStream("file2.xlsx", FileMode.Open));
        var sha3 = SHA256.HashData(new FileStream("file3.xlsx", FileMode.Open));
        Assert.Equal(sha1, sha2);
        Assert.Equal(sha1, sha3);
        Assert.Equal(sha2, sha3);
    }
    
    [Fact]
    public void XlsbWriteT2()
    {
        var dt = GetDataTable();
        {
            using var excel = new XlsbWriter("file1.xlsb");
            excel.AddSheet("sheetName");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX1");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX2");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
        }
        {
            using var fs = File.Open("file2.xlsb", System.IO.FileMode.Create);
            var excel = new XlsbWriter(fs);
            excel.AddSheet("sheetName");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX1");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX2");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.Dispose();
        }
        {
            using var memoryStream = new MemoryStream();
            var excel = new XlsbWriter(memoryStream);
            excel.AddSheet("sheetName");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX1");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.AddSheet("sheetNameX2");
            excel.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
            excel.Dispose();

            using var fileStream = File.Open("file3.xlsb", FileMode.Create);
            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.CopyTo(fileStream);
        }


        var str1 = ReadFileAndCompare("file1.xlsb");
        var str2 = ReadFileAndCompare("file2.xlsb");
        var str3 = ReadFileAndCompare("file3.xlsb");
        Assert.Equal(str1, str2);
        Assert.Equal(str1, str3);
        Assert.Equal(str2, str3);


        var sha1 = SHA256.HashData(new FileStream("file1.xlsb", FileMode.Open));
        var sha2 = SHA256.HashData(new FileStream("file2.xlsb", FileMode.Open));
        var sha3 = SHA256.HashData(new FileStream("file3.xlsb", FileMode.Open));

        //????
        //Assert.Equal(sha1, sha2);
        //Assert.Equal(sha1, sha3);
        //Assert.Equal(sha2, sha3);
    }


    private static string ReadFileAndCompare(string path)
    {
        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        var sheetNames = excelFile.GetScheetNames();
        excelFile.ActualSheetName = sheetNames[0];
        object[]? row = null;
        StringBuilder sb = new StringBuilder();
        while (excelFile.Read())
        {
            row ??= new object[excelFile.FieldCount];
            excelFile.GetValues(row);
            sb.AppendLine(string.Join('|', row));
        }


        return sb.ToString().Trim();
    }

    /// <summary>
    /// get sample data
    /// </summary>
    /// <param name="rowsCount"></param>
    /// <returns></returns>
    private static DataTable GetDataTable(int rowsCount = 100)
    {
        DataTable dataTable = new DataTable();
        dataTable.Columns.Add("COL1_INT", typeof(int));
        dataTable.Columns.Add("COL2_TXT", typeof(string));
        dataTable.Columns.Add("COL3_DATETIME", typeof(DateTime));
        dataTable.Columns.Add("COL4_DOUBLE", typeof(double));

        Random rn = new Random();

        for (int i = 0; i < rowsCount; i++)
        {
            dataTable.Rows.Add(new object[]
            {
                    rn.Next(0, 1_000_000),
                    "TXT_" + rn.Next(0, 1_000_000),
                    new DateTime(rn.Next(1900,2100), rn.Next(1,12), rn.Next(1, 28)),
                    rn.NextDouble()
            });
        }
        return dataTable;
    }


}