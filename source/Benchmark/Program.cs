using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using SpreadSheetTasks;

namespace Benchmark
{
    class Program
    {
        static void Main(string[] args)
        {
#if RELEASE
            //var summary = BenchmarkRunner.Run<ReadBenchXlsx>();
            //var summary = BenchmarkRunner.Run<ReadBenchXlsb>();
            var summary = BenchmarkRunner.Run<WriteBench>();
#endif
#if DEBUG
            Stopwatch st = new Stopwatch();
            st.Start();
            var xlsb = new ReadBenchXlsb();
            xlsb.FileName = "200kFile.xlsb";
            xlsb.ReadFile();
            Console.WriteLine("ReadBenchXlsb: " + st.ElapsedMilliseconds + " ms");

            //var xlsx = new ReadBenchXlsx();
            //xlsx.FileName = "200kFile.xlsx";
            //st.Restart();
            //xlsx.ReadFile();
            //Console.WriteLine("ReadBenchXlsx: " + st.ElapsedMilliseconds + " ms");

            ////write
            //var readerXlsx = new WriteBench();
            //readerXlsx.Rows = 200_000;
            //readerXlsx.setup();

            //st.Restart();
            //readerXlsx.XlsxTestDefault();
            //Console.WriteLine("XlsxTestDefault: " + st.ElapsedMilliseconds + " ms");

            //st.Restart();
            //readerXlsx.XlsxTestLowMemory();
            //Console.WriteLine("XlsxTestLowMemory: " + st.ElapsedMilliseconds + " ms");

            //st.Restart();
            //readerXlsx.XlsbTestDefault();
            //Console.WriteLine("XlsbTestDefault: " + st.ElapsedMilliseconds + " ms");
#endif

        }
    }

    [MemoryDiagnoser]
    public class ReadBenchXlsx
    {

        [Params("200kFile.xlsx")]
        public string FileName { get; set; }

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

    [MemoryDiagnoser]
    public class ReadBenchXlsb
    {

        [Params("200kFile.xlsb")]
        public string FileName { get; set; }

        [Params(false,true)]
        public bool UseMemoryStreamInXlsb { get; set; }

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
                excelFile.UseMemoryStreamInXlsb = UseMemoryStreamInXlsb;
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

    [MemoryDiagnoser]
    public class WriteBench
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
        public void XlsxTestDefault()
        {
            using (XlsxWriter xlsx = new XlsxWriter("fileLowMemory.xlsx"))
            {
                xlsx.AddSheet("sheetName");
                xlsx.WriteSheet(dt.CreateDataReader());
            }
        }

        [Benchmark]
        public void XlsxTestLowMemory()
        {
            using (XlsxWriter xlsx = new XlsxWriter("file.xlsx", bufferSize: 4096, InMemoryMode: false, useScharedStrings: false))
            {
                xlsx.AddSheet("sheetName");
                xlsx.WriteSheet(dt.CreateDataReader());
            }
        }

        [Benchmark]
        public void XlsbTestDefault()
        {
            using (XlsbWriter xlsx = new XlsbWriter("file.xlsb"))
            {
                xlsx.AddSheet("sheetName");
                xlsx.WriteSheet(dt.CreateDataReader());
            }
        }


    }

}
