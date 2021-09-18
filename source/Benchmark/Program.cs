using System;
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
            var summary = BenchmarkRunner.Run<ReadBench>();
#endif
#if DEBUG
            var b1 = new ReadBench();
            b1.FileName = "200kFile.xlsb";
            Stopwatch st = new Stopwatch();
            st.Start();
            b1.ReadFile();
            Console.WriteLine(st.ElapsedMilliseconds);
#endif

        }
    }

    [MemoryDiagnoser]
    public class ReadBench
    {

        [Params("200kFile.xlsx", "200kFile.xlsb")]
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
}
