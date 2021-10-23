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

using SylvanCsv = Sylvan.Data.Csv;
using Sylvan.Data.Excel;

namespace Benchmark
{
    class Program
    {
        static void Main(string[] args)
        {
#if RELEASE
            //var summary = BenchmarkRunner.Run<ReadBenchXlsx>();
            //var summary = BenchmarkRunner.Run<ReadBenchXlsb>();
            var summary = BenchmarkRunner.Run<WriteBenchExcel>();
            //var summary = BenchmarkRunner.Run<CsvReadBench>();
            //var summary = BenchmarkRunner.Run<CsvWriterBench>();

#endif
#if DEBUG
            ExcelTest();
            //CsvTest();
            //CsvWriterTest();
#endif
        }
        static void ExcelTest()
        {
            //ReadBenchXlsb e = new ReadBenchXlsb();

            //e.FileName = "200kFile.xlsb";
            //e.ReadFile();
            WriteBenchExcel e = new WriteBenchExcel();
            e.Rows = 200_000;
            e.setup();
            //e.XlsbWriteDefault();
            e.XlsxWriteDefault();
        }

        static void CsvTest()
        {
            string path = @"C:\sqls\CsvReader\simpleCsv.txt";
            //string path = @"C:\sqls\CsvReader\simpleCsvBig.txt";

            //string path = @"C:\sqls\CsvReader\annual-enterprise-survey-2020-financial-year-provisional-csv.csv";
            //var a = File.ReadAllText(path).Replace("\r\n", "\n");
            //File.WriteAllText(@"C:\sqls\CsvReader\simpleCsvLF.txt", a);
            //Span<char> buff = stackalloc char[256];

            using CsvTextReader rd = new CsvTextReader(path);
            //rd.UseIntrinsic = false;

            while (rd.Read())
            {
                Console.WriteLine("row " + rd.RecordsAffected);
                //Console.WriteLine("row " + ++j);
                //rd.RetrieveStringRow(); tylko dla GetString
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    //Console.WriteLine($"    col {l + 1}: {System.Text.Encoding.UTF8.GetString(rd.GetByteSpan(l))}");
                    Console.WriteLine($"    col {l + 1}: {rd.GetString(l)}");
                    //Console.WriteLine($"    col {l + 1}: {rd.GetSpan(l).ToString()}");
                }
            }

            //Console.WriteLine("records " + rd.RecordsAffected);
            //Console.WriteLine(rd.ss);
        }

        static void CsvWriterTest()
        {
            CsvWriterBench csvWriterBench = new CsvWriterBench();
            csvWriterBench.Rows = 50_000;
            csvWriterBench.setup();
            csvWriterBench.CsvWriterTestA();
            //csvWriterBench.CsvWriterSylvan();
        }
    }

    [SimpleJob(RuntimeMoniker.Net50)]
    [SimpleJob(RuntimeMoniker.Net60)]
    [MemoryDiagnoser]
    public class ReadBenchXlsx
    {

        string filename65k = "65K_Records_Data.xlsx";
        string filename200k = "200kFile.xlsx";

        [Benchmark]
        public void SpreadSheetTasks200K()
        {
            
#if RELEASE
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\..\\..\\..\\..\\FilesToTest\\{filename200k}");
#endif
#if DEBUG
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\FilesToTest\\{filename200k}");
#endif
            using (XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit())
            {
                excelFile.Open(path);
                var sheetNames = excelFile.GetScheetNames();
                excelFile.ActualSheetName = sheetNames[0];
                object[] row = null;
                int rowNum = 0;
                while (excelFile.Read())
                {
                    if (row == null)
                    {
                        row = new object[excelFile.FieldCount];
                    }
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
#if RELEASE
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\..\\..\\..\\..\\FilesToTest\\{filename200k}");
#endif
#if DEBUG
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\FilesToTest\\{filename200k}");
#endif

            var reader = ExcelDataReader.Create(path);

            object[] row = new object[reader.FieldCount];

            while (reader.Read())
            {
                reader.GetValues(row);
            }
        }

        //[Benchmark]
        public void SpreadSheetTasks65k()
        {
#if RELEASE
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\..\\..\\..\\..\\FilesToTest\\{filename65k}");
#endif
#if DEBUG
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\FilesToTest\\{filename65k}");
#endif
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

        //[Benchmark]
        public void Sylvan65K()
        {
#if RELEASE
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\..\\..\\..\\..\\FilesToTest\\{filename65k}");
#endif
#if DEBUG
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\FilesToTest\\{filename65k}");
#endif

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
            var id = excelFile.GetInt64(6);
            var shipDate = excelFile.GetDateTime(7);
            var unitsSold = excelFile.GetInt64(8);
            var unitPrice = excelFile.GetDouble(9);
            var unitCost = excelFile.GetDouble(10);
            var totalRevenue = excelFile.GetDouble(11);
            var totalCost = excelFile.GetDouble(12);
            var totalProfit = excelFile.GetDouble(13);
        }

    }


    [SimpleJob(RuntimeMoniker.Net50)]
    [SimpleJob(RuntimeMoniker.Net60)]
    [MemoryDiagnoser]
    public class ReadBenchXlsb
    {

        [Params("200kFile.xlsb")]
        public string FileName { get; set; }

        [Params(false,true)]
        public bool InMemory { get; set; }

        [Benchmark]
        public void ReadFile()
        {
#if RELEASE
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\..\\..\\..\\..\\FilesToTest\\{FileName}");
#endif
#if DEBUG
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"..\\..\\..\\FilesToTest\\{FileName}");
#endif
            using (XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit())
            {
                excelFile.UseMemoryStreamInXlsb = InMemory;
                excelFile.Open(path);
                excelFile.ActualSheetName = "sheet1";
                object[] row = null;
                int rowNum = 0;
                while (excelFile.Read())
                {
                    if (row == null)
                    {
                        row = new object[excelFile.FieldCount];
                    }
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


    [SimpleJob(RuntimeMoniker.Net50)]
    [SimpleJob(RuntimeMoniker.Net60)]
    [MemoryDiagnoser]
    public class WriteBenchExcel
    {

        [Params(200_000)]
        public int Rows { get; set; }

        DataTable dt = new DataTable();

        [GlobalSetup]
        public void setup()
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


        [Benchmark]
        public void XlsxWriteDefault()
        {
            //using (XlsxWriterTest xlsx = new XlsxWriterTest(@"C:\sqls\fileLowMemory.xlsx"))
            using (XlsxWriter xlsx = new XlsxWriter("fileLowMemory.xlsx"))
            {
                xlsx.AddSheet("sheetName");
                xlsx.WriteSheet(dt.CreateDataReader());
            }
        }

        [Benchmark]
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


    [SimpleJob(RuntimeMoniker.Net50)]
    [SimpleJob(RuntimeMoniker.Net60)]
    [MemoryDiagnoser]
    public class CsvReadBench
    {
        //string path = @"C:\sqls\CsvReader\1000000 Sales Records.csv";
        string path = @"C:\sqls\CsvReader\annual-enterprise-survey-2020-financial-year-provisional-csv.csv";

        [Benchmark]
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

        [Benchmark]
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

        [Benchmark]
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

        [Benchmark]
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

    [SimpleJob(RuntimeMoniker.Net50)]
    [SimpleJob(RuntimeMoniker.Net60)]
    [MemoryDiagnoser]
    public class CsvWriterBench
    {
        [Params(500_000)]
        public int Rows { get; set; }

        public int BufferSize { get; set; }

        DataTable dt = new DataTable();

        string path = @"C:\sqls\testWriter.txt";

        [GlobalSetup]
        public void setup()
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

        [Benchmark]
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
                DateFormat = "yyyy-MM-dd HH:mm:ss"
            };
            using var cw = SylvanCsv.CsvDataWriter.Create(path);
            cw.Write(dt.CreateDataReader());
        }

    }
}
