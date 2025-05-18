using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using SpreadSheetTasks;
using Sylvan.Data.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO.Compression;


namespace Benchmark
{
    class Program
    {
        static void Main(string[] args)
        {
#if RELEASE
            //var summary = BenchmarkRunner.Run<ReadBenchXlsx>();
            var summary2 = BenchmarkRunner.Run<ReadBenchXlsb>();
            //var summary3 = BenchmarkRunner.Run<WriteBenchExcel>();
#endif
#if DEBUG
            Console.WriteLine("PLEASE RUN IN RELEASE MODE");
#endif
        }
    }
    [SimpleJob(RuntimeMoniker.Net90)]
    [MemoryDiagnoser]
    public class ReadBenchXlsx
    {
        private readonly string _baseFilePath = @"D:/DEV/source/repos/PublicNuget/SpreadSheetTasks";
        readonly string filename65k = "65K_Records_Data.xlsx";
        readonly string filename200k = "200kFile.xlsx";

        [Benchmark]
        public void SpreadSheetTasks200K()
        {
            var path = $@"{_baseFilePath}/source/Benchmark/FilesToTest/{filename200k}";

            using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
            excelFile.Open(path);
            var sheetNames = excelFile.GetScheetNames();
            excelFile.ActualSheetName = sheetNames[0];
            object[] row = null;
            int rowNum = 0;
            while (excelFile.Read())
            {
                row ??= new object[excelFile.FieldCount];
                excelFile.GetValues(row);
                rowNum++;
#if DEBUG
                    if (rowNum % 10_000 == 0)
                        Console.WriteLine("row " + rowNum.ToString("N0") + ": " + String.Join('|', row));
#endif
            }
        }

        [Benchmark]
        public void Sylvan200k()
        {
            var path = @$"{_baseFilePath}/source/Benchmark/FilesToTest/{filename200k}";

            var reader = ExcelDataReader.Create(path);

            object[] row = new object[reader.FieldCount];

            while (reader.Read())
            {
                reader.GetValues(row);
            }
        }

        [Benchmark]
        public void SpreadSheetTasks65k()
        {
            var path = @$"{_baseFilePath}/source/Benchmark/FilesToTest/{filename65k}";

            using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
            excelFile.Open(path);
            var sheetNames = excelFile.GetScheetNames();
            excelFile.ActualSheetName = sheetNames[0];

            excelFile.Read(); // = skip header
            while (excelFile.Read())
            {
                ProcessRecord(excelFile);
            }
        }

        [Benchmark]
        public void Sylvan65K()
        {
            var path = $@"{_baseFilePath}/source/Benchmark/FilesToTest/{filename65k}";

            var reader = Sylvan.Data.Excel.ExcelDataReader.Create(path);

            do
            {
                while (reader.Read())
                {
                    ProcessRecordSylvan(reader);
                }

            } while (reader.NextResult());
        }

        //method from
        //https://github.com/MarkPflug/Benchmarks/blob/main/source/Benchmarks/XlsxDataReaderBenchmarks.cs
        static void ProcessRecordSylvan(IDataReader reader)
        {
            var region = reader.GetString(0);
            var country = reader.GetString(1);
            var type = reader.GetString(2);
            var channel = reader.GetString(3);
            var priority = reader.GetString(4);
            var orderDate = reader.GetDateTime(5);
            var id = reader.GetInt32(6);
            var shipDate = reader.GetDateTime(7);
            var unitsSold = reader.GetInt32(8);
            var unitPrice = reader.GetDouble(9);
            var unitCost = reader.GetDouble(10);
            var totalRevenue = reader.GetDouble(11);
            var totalCost = reader.GetDouble(12);
            var totalProfit = reader.GetDouble(13);
        }

        static void ProcessRecord(XlsxOrXlsbReadOrEdit excelFile)
        {
            var region = excelFile.GetString(0);
            var country = excelFile.GetString(1);
            var type = excelFile.GetString(2);
            var channel = excelFile.GetString(3);
            var priority = excelFile.GetString(4);
            var orderDate = excelFile.GetDateTime(5);
            var id = excelFile.GetInt32(6);
            var shipDate = excelFile.GetDateTime(7);
            var unitsSold = excelFile.GetInt32(8);
            var unitPrice = excelFile.GetDouble(9);
            var unitCost = excelFile.GetDouble(10);
            var totalRevenue = excelFile.GetDouble(11);
            var totalCost = excelFile.GetDouble(12);
            var totalProfit = excelFile.GetDouble(13);
        }
    }

    [SimpleJob(RuntimeMoniker.Net90)]
    [MemoryDiagnoser]
    public class ReadBenchXlsb
    {
        private readonly string _baseFilePath = @"D:/DEV/source/repos/PublicNuget/SpreadSheetTasks";
        [Params("65K_Records_Data.xlsb", "200kFile.xlsb")]
        public string FileName { get; set; }

        [Benchmark(Description = "SpreadSheetTasks - XLSB Read - v1")]
        public void ReadFile()
        {
            var path = $@"{_baseFilePath}/source/Benchmark/FilesToTest/{FileName}";
            using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
            excelFile.Open(path);
            excelFile.ActualSheetName = "sheet1";
            object[] row = null;
            int rowNum = 0;
            while (excelFile.Read())
            {
                row ??= new object[excelFile.FieldCount];
                excelFile.GetValues(row);
                rowNum++;
            }
        }

        [Benchmark(Description = "SpreadSheetTasks - XLSB Read - v2")]
        public void ReadFileInMemory()
        {
            var path = $@"{_baseFilePath}/source/Benchmark/FilesToTest/{FileName}";
            using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
            excelFile.Open(path);
            excelFile.UseMemoryStreamInXlsb = false;
            excelFile.ActualSheetName = "sheet1";
            object[] row = null;
            int rowNum = 0;
            while (excelFile.Read())
            {
                row ??= new object[excelFile.FieldCount];
                excelFile.GetValues(row);
                rowNum++;
            }
        }

        [Benchmark(Description = "Sylvan.Data.Excel - XLSB Read")]
        public void Sylvan()
        {
            var path = $@"{_baseFilePath}/source/Benchmark/FilesToTest/{FileName}";
            using ExcelDataReader reader = ExcelDataReader.Create(path);
            object[] row = new object[reader.FieldCount];
            while (reader.Read())
            {
                reader.GetValues(row);
            }
        }
    }

    [SimpleJob(RuntimeMoniker.Net90)]
    [MemoryDiagnoser]
    public class WriteBenchExcel
    {
        public int RowsCount = 200_000;
        private readonly Dictionary<string,DataTable> dataTables = new Dictionary<string, DataTable>
        {
            {"GENERAL", new DataTable() },
            {"INT", new DataTable() },
            {"DOUBLE", new DataTable() },
            {"DATETIME", new DataTable() },
            {"TEXT", new DataTable() },
        };

        //[Params("GENERAL", "INT", "DOUBLE", "DATETIME", "TEXT")]
        [Params("GENERAL")]
        public string ReaderType { get; set; } = "GENERAL";
        private DataTable Dt => dataTables[ReaderType];

        [GlobalSetup]
        public void Setup()
        {
            dataTables["GENERAL"].Columns.Add("COL1_INT", typeof(int));
            dataTables["GENERAL"].Columns.Add("COL2_TXT", typeof(string));
            dataTables["GENERAL"].Columns.Add("COL3_DATETIME", typeof(DateTime));
            dataTables["GENERAL"].Columns.Add("COL4_DOUBLE", typeof(double));

            Random rn = new Random();

            for (int i = 0; i < RowsCount; i++)
            {
                dataTables["GENERAL"].Rows.Add(new object[]
                {
                    rn.Next(0, 1_000_000),
                    "TXT_" + rn.Next(0, 1_000_000),
                    new DateTime(rn.Next(1900,2100), rn.Next(1,12), rn.Next(1, 28)),
                    rn.NextDouble()
                });
            }

            dataTables["INT"].Columns.Add("COL1", typeof(int));
            dataTables["INT"].Columns.Add("COL2", typeof(int));
            dataTables["INT"].Columns.Add("COL3", typeof(int));
            dataTables["INT"].Columns.Add("COL4", typeof(int));

            for (int i = 0; i < RowsCount; i++)
            {
                dataTables["INT"].Rows.Add(new object[]
                {
                    rn.Next(),
                    rn.Next(),
                    rn.Next(),
                    rn.Next(),
                });
            }


            dataTables["DOUBLE"].Columns.Add("COL1", typeof(double));
            dataTables["DOUBLE"].Columns.Add("COL2", typeof(double));
            dataTables["DOUBLE"].Columns.Add("COL3", typeof(double));
            dataTables["DOUBLE"].Columns.Add("COL4", typeof(double));

            for (int i = 0; i < RowsCount; i++)
            {
                dataTables["DOUBLE"].Rows.Add(new object[]
                {
                    rn.NextDouble(),
                    rn.NextDouble(),
                    rn.NextDouble(),
                    rn.NextDouble(),
                });
            }

            dataTables["DATETIME"].Columns.Add("COL1", typeof(DateTime));
            dataTables["DATETIME"].Columns.Add("COL2", typeof(DateTime));
            dataTables["DATETIME"].Columns.Add("COL3", typeof(DateTime));
            dataTables["DATETIME"].Columns.Add("COL4", typeof(DateTime));

            for (int i = 0; i < RowsCount; i++)
            {
                dataTables["DATETIME"].Rows.Add(new object[]
                {
                    DateTime.Now.AddSeconds(i),
                    DateTime.Now.AddSeconds(i+1),
                    DateTime.Now.AddSeconds(i+2),
                    DateTime.Now.AddSeconds(i+3),
                });
            }

            dataTables["TEXT"].Columns.Add("COL1", typeof(string));
            dataTables["TEXT"].Columns.Add("COL2", typeof(string));
            dataTables["TEXT"].Columns.Add("COL3", typeof(string));
            dataTables["TEXT"].Columns.Add("COL4", typeof(string));

            for (int i = 0; i < RowsCount; i++)
            {
                dataTables["TEXT"].Rows.Add(new object[]
                {
                    "TXT_" + rn.Next(0, 1_000_000),
                    "TXT_" + rn.Next(0, 1_000_000),
                    "TXT_" + rn.Next(0, 1_000_000),
                    "TXT_" + rn.Next(0, 1_000_000),
                });
            }
        }

        private readonly CompressionLevel _cLvl = CompressionLevel.Fastest;

        [Benchmark(Description = "SpreadSheetTasks - XLSB Write")]
        public void XlsbWriteDefault()
        {
            using XlsbWriter xlsx = new XlsbWriter("file.xlsb", _cLvl);
            xlsx.AddSheet("sheetName");
            xlsx.WriteSheet(Dt.CreateDataReader(), doAutofilter: true);
        }

        [Benchmark]
        public void XlsbSylvanWrite()
        {
            using var edw = ExcelDataWriter.Create("fileSylvan.xlsb", new ExcelDataWriterOptions() {  CompressionLevel = _cLvl });
            DbDataReader dr;
            dr = Dt.CreateDataReader();
            edw.Write(dr, "sheetName");
        }


        [Benchmark(Description = "SpreadSheetTasks - XLSX Write")]
        public void XlsxWriteLowMemory()
        {
            using XlsxWriter xlsx = new XlsxWriter("file.xlsx");
            xlsx.AddSheet("sheetName");
            xlsx.WriteSheet(Dt.CreateDataReader());
        }

        [Benchmark(Description = "Sylvan - XLSX Write")]
        public void XlsxSylvanWrite()
        {
            using var edw = ExcelDataWriter.Create("fileSylvan.xlsx");
            DbDataReader dr;
            dr = Dt.CreateDataReader();
            edw.Write(dr, "sheetName");
        }

    }
 }
