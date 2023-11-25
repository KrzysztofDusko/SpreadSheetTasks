using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO.Compression;
using System.Numerics;
using System.Runtime.CompilerServices;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using CommunityToolkit.HighPerformance.Buffers;
using nietras.SeparatedValues;
using SpreadSheetTasks;
using SpreadSheetTasks.CsvWriter;
using Sylvan.Data.Excel;
using SylvanCsv = Sylvan.Data.Csv;


namespace Benchmark
{
    class Program
    {
        static void Main(string[] args)
        {
#if RELEASE
            //var summary = BenchmarkRunner.Run<ReadBenchXlsx>();
            //var summary2 = BenchmarkRunner.Run<ReadBenchXlsb>();
            var summary3 = BenchmarkRunner.Run<WriteBenchExcel>();
            //var summary4 = BenchmarkRunner.Run<CsvReadBench>();
            //var summary5 = BenchmarkRunner.Run<CsvWriterBench>();
            //var sumary = BenchmarkRunner.Run<NumberParseTest>();
            //var summary7 = BenchmarkRunner.Run<StringPoolSylvanTest>();
#endif
#if DEBUG
            var b = new WriteBenchExcel();
            b.Setup();
            b.ReaderType = "INT";
            b.XlsbWriteDefault();
            //b.XlsbSylvanWrite();
            //micro.DoJob();
            //micro.DoM4();
#endif
        }
    }


    [SimpleJob(RuntimeMoniker.Net80)]
    [MemoryDiagnoser]
    public class ReadBenchXlsx
    {
        readonly string filename65k = "65K_Records_Data.xlsx";
        readonly string filename200k = "200kFile.xlsx";

        //[Benchmark]
        public void SpreadSheetTasks200K()
        {
            var path = $@"E:\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{filename200k}";

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

        //[Benchmark]
        public void Sylvan200k()
        {
            var path = @$"E:\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{filename200k}";

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
            var path = @$"E:\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{filename65k}";

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
            var path = $@"E:\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{filename65k}";

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


    [SimpleJob(RuntimeMoniker.Net80)]
    [MemoryDiagnoser]
    public class ReadBenchXlsb
    {
        [Params("65K_Records_Data.xlsb", "200kFile.xlsb")]
        public string FileName { get; set; }

        [Benchmark(Description = "SpreadSheetTasks - XLSB Read - v1")]
        public void ReadFile()
        {
            var path = $@"E:\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{FileName}";
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
            var path = $@"E:\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{FileName}";
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
            var path = $@"E:\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{FileName}";
            using ExcelDataReader reader = ExcelDataReader.Create(path);
            object[] row = new object[reader.FieldCount];
            while (reader.Read())
            {
                reader.GetValues(row);
            }
        }
    }


    [SimpleJob(RuntimeMoniker.Net80)]
    [MemoryDiagnoser]
    public class WriteBenchExcel
    {

        private const int RowsCount = 200_000;
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

        System.IO.Compression.CompressionLevel cLvl = System.IO.Compression.CompressionLevel.Fastest;

        [Benchmark(Description = "SpreadSheetTasks - XLSB Write")]
        public void XlsbWriteDefault()
        {
            using XlsbWriter xlsx = new XlsbWriter("file.xlsb", cLvl);
            xlsx.AddSheet("sheetName");
            xlsx.WriteSheet(Dt.CreateDataReader(), doAutofilter: true);
        }

        [Benchmark]
        public void XlsbSylvanWrite()
        {
            using var edw = ExcelDataWriter.Create("fileSylvan.xlsb", new ExcelDataWriterOptions() {  CompressionLevel = cLvl});
            DbDataReader dr;
            dr = Dt.CreateDataReader();
            edw.Write(dr, "sheetName");
        }


        //[Benchmark(Description = "SpreadSheetTasks - XLSX Write - v2")]
        public void XlsxWriteLowMemory()
        {
            using XlsxWriter xlsx = new XlsxWriter("file.xlsx", bufferSize: 4096, InMemoryMode: false, useScharedStrings: false);
            xlsx.AddSheet("sheetName");
            xlsx.WriteSheet(Dt.CreateDataReader());
        }

        //[Benchmark(Description = "SpreadSheetTasks - XLSX Write - v2")]
        public void XlsxWriteDefault()
        {
            using XlsxWriter xlsx = new XlsxWriter("fileLowMemory.xlsx");
            xlsx.AddSheet("sheetName");
            xlsx.WriteSheet(Dt.CreateDataReader(), doAutofilter: true);
        }

    }

    [SimpleJob(RuntimeMoniker.Net80)]
    [MemoryDiagnoser]
    public class CsvReadBench
    {
        readonly string path = @$"E:\sqls\CsvReader\annual-enterprise-survey-2020-financial-year-provisional-csv.csv";
        int N = 20;


        [Benchmark]
        public void SylvanSPAN()
        {
            for (int i = 0; i < N; i++)
            {
                var rd = SylvanCsv.CsvDataReader.Create(path/*, opt*/);

                while (rd.Read())
                {
                    for (int l = 0; l < rd.FieldCount; l++)
                    {
                        var x = rd.GetFieldSpan(l);
                    }
                }
            }
        }
        [Benchmark]
        public void SepSPAN()
        {
            for (int i = 0; i < N; i++)
            {
                using var rd = Sep.Reader().FromFile(path);
                foreach (var readRow in rd)
                {
                    for (int l = 0; l < readRow.ColCount; l++)
                    {
                        var x = readRow[l].Span;
                    }
                }
            }
        }
        [Benchmark]
        public void SylvanString()
        {
            for (int i = 0; i < N; i++)
            {
                var rd = SylvanCsv.CsvDataReader.Create(path/*, opt*/);

                while (rd.Read())
                {
                    for (int l = 0; l < rd.FieldCount; l++)
                    {
                        rd.GetString(l);
                    }
                }
            }
        }
        [Benchmark]
        public void SepGetString()
        {
            for (int i = 0; i < N; i++)
            {
                using var rd = Sep.Reader().FromFile(path);
                foreach (var readRow in rd)
                {
                    for (int l = 0; l < readRow.ColCount; l++)
                    {
                        readRow[l].ToString();
                    }
                }
            }
        }

        //[Benchmark]
        public void SylvanSpanString()
        {
            List<string> list = new List<string>();
            SimpleStringPool stringPool = new SimpleStringPool();
            for (int i = 0; i < N; i++)
            {
                var rd = SylvanCsv.CsvDataReader.Create(path/*, opt*/);

                while (rd.Read())
                {
                    for (int l = 0; l < rd.FieldCount; l++)
                    {
                        list.Add(stringPool.GetString(rd.GetFieldSpan(l)));
                    }
                }
            }
        }

        //[Benchmark]
        public void SepAllColumnsSpanString()
        {
            List<string> list = new List<string>();
            for (int i = 0; i < N; i++)
            {
                using var rd = Sep.Reader(o => o with
                {
                    CreateToString = SepToString.OnePool(),
                }).FromFile(path);
                foreach (var readRow in rd)
                {
                    for (int l = 0; l < readRow.ColCount; l++)
                    {
                        list.Add(readRow[l].ToString());
                    }
                }
            }
        }


        //[Benchmark]
        public void SepOneColumnSpanString()
        {
            List<string> list = new List<string>();
            for (int i = 0; i < N; i++)
            {
                using var rd = Sep.Reader(o => o with
                {
                    CreateToString = SepToString.PoolPerCol(),
                }).FromFile(path);
                foreach (var readRow in rd)
                {
                    for (int l = 0; l < readRow.ColCount; l++)
                    {
                        list.Add(readRow[l].ToString());
                    }
                }
            }
        }

        sealed class MyPool : SepToString
        {
            public static readonly MyPool Instance = new MyPool();

            private static readonly SimpleStringPool stringPool = new SimpleStringPool();

            public override string ToString(ReadOnlySpan<char> colSpan, int colIndex)
            {
                return stringPool.GetString(colSpan);
            }
        }

    }

    [SimpleJob(RuntimeMoniker.Net80)]
    [MemoryDiagnoser]
    public class CsvWriterBench
    {
        [Params(500_000)]
        public int Rows { get; set; }

        public int BufferSize { get; set; }

        readonly DataTable dt = new DataTable();
        readonly string path = @$"E:\sqls\CsvReader\testWriter.txt";

        [GlobalSetup]
        public void Setup()
        {
            dt.Columns.Add("COL1_INT", typeof(int));
            dt.Columns.Add("COL2_TXT", typeof(string));
            dt.Columns.Add("COL3_DATETIME", typeof(DateTime));
            dt.Columns.Add("COL4_DOUBLE", typeof(double));

            Random rn = new Random();

            for (int i = 0; i < Rows; i++)
            {
                dt.Rows.Add(new object[]
                {
                    i == Rows/2 ? DBNull.Value:rn.Next(1,10_000)
                    , i == Rows/2 ? DBNull.Value:"TXT|_" + rn.Next(1,10_000)
                    , i == Rows/2 ? DBNull.Value:new DateTime(rn.Next(1900,2100), rn.Next(1,12), rn.Next(1, 28))
                    , i == Rows/2 ? DBNull.Value:rn.NextDouble()
                });
            }
        }

        //[Benchmark]
        public void CsvWriterTestA()
        {
            CsvWriter cw = new CsvWriter(path);
            cw.Write(dt.CreateDataReader());
        }


        [Benchmark]
        public void CsvWriterSylvan()
        {
            SylvanCsv.CsvDataWriterOptions opt = new SylvanCsv.CsvDataWriterOptions()
            {
                DateTimeFormat = "yyyy-MM-dd HH:mm:ss"
            };
            using var cw = SylvanCsv.CsvDataWriter.Create(path);
            cw.Write(dt.CreateDataReader());
        }

    }

    [SimpleJob(RuntimeMoniker.Net80)]
    public class NumberParseTest
    {

        [Params(2,4,8,16)]
        public int len { get; set; } = 16;

        //int len = 16;
        char[] buff = new char[] { '1', '2', '5', '9', '2', '5', '9', '9', '1', '2', '5', '9', '2', '5', '9', '9' };

        //[Benchmark]
        public Int64 ParseToInt64FromBuffer()
        {
            Int64 res = 0;
            int start = buff[0] == '-' ? 1 : 0;
            for (int i = start; i < len; i++)
            {
                res = res * 10 + (buff[i] - '0');
            }
            return start == 1 ? -res : res;
        }


        //[Benchmark]
        public Int64 ParseToInt64FromBuffer2()
        {
            int start = buff[0] == '-' ? 1 : 0;
            Int64 res = buff[start] - '0';
            for (int i = start + 1; i < len; i++)
            {
                res = res * 10 + (buff[i] - '0');
            }
            return start == 1 ? -res : res;
        }

        //[Benchmark]
        public Int64 ParseToInt64FromBuffer3()
        {
            var c1 = buff[0] == '-';
            byte start = Unsafe.As<bool, byte>(ref c1);
            Int64 res = buff[start] - '0';
            for (int i = start + 1; i < len; i++)
            {
                res = res * 10 + (buff[i] - '0');
            }
            //return res * (1 - 2*start);
            return start == 1 ? -res : res;
        }

        //[Benchmark]
        public Int64 ParseToInt64System()
        {
            return Int64.Parse(buff.AsSpan(), System.Globalization.NumberStyles.AllowLeadingSign | System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture);
        }

        [Benchmark]
        public bool ContainsDoubleMarks()
        {
            for (int i = 0; i < len; i++)
            {
                char c = buff[i];
                if (c == '.' || c == 'E')
                {
                    return true;
                }
            }
            return false;
        }


        [Benchmark]
        public bool ContainsDoubleMarks2()
        {
            return buff.AsSpan(0, len).IndexOfAny('.','E') > 0;
        }
    }


    [SimpleJob(RuntimeMoniker.Net80)]
    [MemoryDiagnoser]
    public class StringPoolSylvanTest
    {
        List<string> ls = new List<string>();
        [GlobalSetup]
        public void FillList()
        {
            Random r = new Random(10);
            for (int i = 0; i < 100_000; i++)
            {
                ls.Add(r.Next().ToString());
            }
        }


        [Params(16,32)]
        public int N { get; set; }

        [Benchmark]
        public void BenchSylvanPoolSylvan()
        {
            SimpleStringPool stringPool = new SimpleStringPool();
            for (int i = 0; i < N; i++)
            {
                foreach (var item in ls)
                {
                    stringPool.GetString(item.AsSpan());
                }
            }
        }

        [Benchmark]
        public void BenchStringPool()
        {
            StringPool stringPool = new StringPool();
            for (int i = 0; i < N; i++)
            {
                foreach (var item in ls)
                {
                    stringPool.GetOrAdd(item.AsSpan());
                }
            }
        }
    }


}
