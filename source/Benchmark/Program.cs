using System;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;

using SpreadSheetTasks;
using SpreadSheetTasks.CsvReader;
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
            var summary2 = BenchmarkRunner.Run<ReadBenchXlsb>();
            //var summary3 = BenchmarkRunner.Run<WriteBenchExcel>();
            //var summary4 = BenchmarkRunner.Run<CsvReadBench>();
            //var summary5 = BenchmarkRunner.Run<CsvWriterBench>();
            //var sumary = BenchmarkRunner.Run<NumberParseTest>();

#endif
#if DEBUG
            var b = new WriteBenchExcel();
            b.Rows = 200_000;
            b.Setup();
            b.XlsbWriteDefault();

            //ExcelTest();
            //CsvTest();
            //CsvWriterTest();
#endif
        }
    }

    //[SimpleJob(RuntimeMoniker.Net60)]
    [SimpleJob(RuntimeMoniker.Net70)]
    //[SimpleJob(RuntimeMoniker.NativeAot70)]
    [MemoryDiagnoser]
    public class ReadBenchXlsx
    {
        readonly string filename65k = "65K_Records_Data.xlsx";
        readonly string filename200k = "200kFile.xlsx";

        //[Benchmark]
        public void SpreadSheetTasks200K()
        {
            var path = $@"C:\Users\dusko\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{filename200k}";

            using (XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit())
            {
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
        }

        //[Benchmark]
        public void Sylvan200k()
        {
            var path = @$"C:\Users\dusko\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{filename200k}";

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
            var path = @$"C:\Users\dusko\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{filename65k}";

            using (XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit())
            {
                excelFile.Open(path);
                var sheetNames = excelFile.GetScheetNames();
                excelFile.ActualSheetName = sheetNames[0];

                excelFile.Read(); // = skip header
                while (excelFile.Read())
                {
                    ProcessRecord(excelFile);
                }
            }
        }

        [Benchmark]
        public void Sylvan65K()
        {
            var path = $@"C:\Users\dusko\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{filename65k}";

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


    //[SimpleJob(RuntimeMoniker.Net60)]
    [SimpleJob(RuntimeMoniker.Net70)]
    //[SimpleJob(RuntimeMoniker.NativeAot70)]
    [MemoryDiagnoser]
    public class ReadBenchXlsb
    {
        [Params("200kFile.xlsb")]
        public string FileName { get; set; }

        [Params(false, true)]
        public bool InMemory { get; set; }

        [Benchmark]
        public void ReadFile()
        {
            var path = $@"C:\Users\dusko\source\repos\SpreadSheetTasks\source\Benchmark\FilesToTest\{FileName}";

            using (XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit())
            {
                excelFile.UseMemoryStreamInXlsb = InMemory;
                excelFile.Open(path);
                excelFile.ActualSheetName = "sheet1";
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
        }
    }


    //[SimpleJob(RuntimeMoniker.Net60)]
    [SimpleJob(RuntimeMoniker.Net70)]
    //[SimpleJob(RuntimeMoniker.NativeAot70)]
    [MemoryDiagnoser]
    public class WriteBenchExcel
    {

        [Params(200_000)]
        public int Rows { get; set; }

        readonly DataTable dt = new DataTable();

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
                    rn.Next(0, 1_000_000)
                    , "TXT_" + rn.Next(0, 1_000_000)
                    , new DateTime(rn.Next(1900,2100), rn.Next(1,12), rn.Next(1, 28))
                    , rn.NextDouble()
                });
            }
        }


        //[Benchmark]
        public void XlsxWriteDefault()
        {
            using (XlsxWriter xlsx = new XlsxWriter("fileLowMemory.xlsx"))
            {
                xlsx.AddSheet("sheetName");
                xlsx.WriteSheet(dt.CreateDataReader());
            }
        }

        //[Benchmark]
        public void XlsxWriteLowMemory()
        {
            using (XlsxWriter xlsx = new XlsxWriter("file.xlsx", bufferSize: 4096, InMemoryMode: false, useScharedStrings: false))
            {
                xlsx.AddSheet("sheetName");
                xlsx.WriteSheet(dt.CreateDataReader());
            }
        }

        [Benchmark]
        public void XlsbWriteDefault()
        {
            using (XlsbWriter xlsx = new XlsbWriter("file.xlsb"))
            {
                xlsx.AddSheet("sheetName");
                xlsx.WriteSheet(dt.CreateDataReader());
            }
        }
    }


    [SimpleJob(RuntimeMoniker.Net60)]
    [SimpleJob(RuntimeMoniker.Net70)]
    [SimpleJob(RuntimeMoniker.NativeAot70)]
    [MemoryDiagnoser]
    public class CsvReadBench
    {
        readonly string path = @$"C:\Users\dusko\sqls\CsvReader\annual-enterprise-survey-2020-financial-year-provisional-csv.csv";

        //[Benchmark]
        public void BinaryReaderGetByteSpan()
        {
            using CsvBinaryReader rd = new CsvBinaryReader(path);

            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetByteSpan(l);
                }
            }
        }

        //[Benchmark]
        public void BinaryReaderGetReadOnlyCharSpan()
        {
            using CsvBinaryReader rd = new CsvBinaryReader(path);
            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetCharSpan(l);
                }
            }
        }

        //[Benchmark]
        public void BinaryReaderGetReadOnlyCharSpanASCII()
        {
            using CsvBinaryReader rd = new CsvBinaryReader(path);
            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetCharSpan(l, System.Text.Encoding.ASCII);
                }
            }
        }

        //[Benchmark]
        public void BinaryReaderGetReadOnlySpanIntrinsicOFF()
        {
            using CsvBinaryReader rd = new CsvBinaryReader(path);
            rd.UseIntrinsic = false;
            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetCharSpan(l);
                }
            }

        }

        //[Benchmark]
        public void BinaryReaderGetStringFromUTF8()
        {
            using CsvBinaryReader rd = new CsvBinaryReader(path);
            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetString(l);
                }
            }
        }

        //[Benchmark]
        public void BinaryReaderGetStringASCII()
        {
            using CsvBinaryReader rd = new CsvBinaryReader(path);
            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetString(l, System.Text.Encoding.ASCII);
                }
            }
        }

        //[Benchmark]
        public void BinaryReaderGetStringASCII2()
        {
            using CsvBinaryReader rd = new CsvBinaryReader(path);
            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    System.Text.Encoding.ASCII.GetString(rd.GetByteSpan(l));
                }
            }
        }

        //[Benchmark]
        public void TextReaderGetSpan()
        {
            using CsvTextReader rd = new CsvTextReader(path);

            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetSpan(l);
                }
            }
        }

        //[Benchmark]
        public void TextReaderGetStringFromSpan()
        {
            using CsvTextReader rd = new CsvTextReader(path);

            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetSpan(l).ToString();
                }
            }
        }

        [Benchmark]
        public void TextReaderGetString()
        {
            using CsvTextReader rd = new CsvTextReader(path);

            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetString(l);
                }
            }
        }

        //[Benchmark]
        public void TextReaderGetStringIgnoreQuoted()
        {
            using CsvTextReader rd = new CsvTextReader(path);

            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetStringIgnoreQuoted(l);
                }
            }
        }

        //[Benchmark]
        public void SylvanString()
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

    [SimpleJob(RuntimeMoniker.Net60)]
    [SimpleJob(RuntimeMoniker.Net70)]
    [SimpleJob(RuntimeMoniker.NativeAot70)]
    [MemoryDiagnoser]
    public class CsvWriterBench
    {
        [Params(500_000)]
        public int Rows { get; set; }

        public int BufferSize { get; set; }

        readonly DataTable dt = new DataTable();
        readonly string path = @$"C:\Users\dusko\sqls\CsvReader\testWriter.txt";

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


    //[SimpleJob(RuntimeMoniker.Net60)]
    [SimpleJob(RuntimeMoniker.Net70)]
    //[SimpleJob(RuntimeMoniker.NativeAot70)]
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

}
