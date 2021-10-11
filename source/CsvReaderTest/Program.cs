using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using Sylvan.Data.Csv;


namespace CsvReaderTest
{
    class Program
    {

        //[SkipLocalsInit]
        static void Main(string[] args)
        {

#if RELEASE
            var summary = BenchmarkRunner.Run<Benchy>();
            return;
#endif

            Stopwatch st = new Stopwatch();
            //Span<char> buff = stackalloc char[256];
            string path = @"C:\sqls\CsvReader\simpleCsv.txt";
            //"C:\sqls\CsvReader\annual-enterprise-survey-2020-financial-year-provisional-csv.csv"
            for (int i = 0; i < 1; i++)
            {
                st.Restart();
                using CsvBinaryReader rd = new CsvBinaryReader(path);

                int j = 0;
                while (rd.Read())
                {
                    Console.WriteLine("row " + rd.RecordsAffected);
                    //Console.WriteLine("row " + ++j);
                    //rd.RetrieveStringRow(); tylko dla GetString
                    for (int l = 0; l < rd.FieldCount; l++)
                    {
                        //Console.WriteLine($"    col {l + 1}: {rd.GetReadOnlySpan(l).ToString()}");
                        Console.WriteLine($"    col {l + 1}: {System.Text.Encoding.UTF8.GetString(rd.GetByteSpan(l))}");
                        //rd.GetString(l);
                        //rd.GetString2(l);
                        //rd.GetReadOnlySpan(l);
                        //rd.GetReadOnlySpan2(l, buff);
                    }
                }

                //Console.WriteLine("records " + rd.RecordsAffected);
                //Console.WriteLine(rd.ss);
   
                Console.WriteLine($"{st.ElapsedMilliseconds}");
            }
        }
    }

    [MemoryDiagnoser]
    public class Benchy
    {
        string path = @"C:\sqls\CsvReader\1000000 Sales Records.csv";

        [Benchmark]
        public void MyGetByteSpan()
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
        //public void MyGetByteSpanOLD()
        //{
        //    using CsvBinaryReaderOld rd = new CsvBinaryReaderOld(path);

        //    while (rd.Read())
        //    {
        //        for (int l = 0; l < rd.FieldCount; l++)
        //        {
        //            rd.GetByteSpan(l);
        //        }
        //    }
        //}


        [Benchmark]
        public void MyGetReadOnlySpan()
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

        [Benchmark]
        [SkipLocalsInit]
        public void MyGetReadOnlySpan2()
        {
            Span<char> buff = stackalloc char[256];
            using CsvBinaryReader rd = new CsvBinaryReader(path);

            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetCharSpanWithBuffer(l, buff);
                }
            }
        }

        //[Benchmark]
        //public void MojeGetStringFromRow()
        //{
        //    using CsvBinaryReader rd = new CsvBinaryReader();
        //    rd.Open(path);
        //    while (rd.Read())
        //    {
        //        rd.RetrieveStringRow();
        //        for (int l = 0; l < rd.FieldCount; l++)
        //        {
        //            rd.GetStringFromRow(l);
        //        }
        //    }
        //}

        [Benchmark]
        public void MyGetString()
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

        [Benchmark]
        public void SylvanString()
        {
            var rd = CsvDataReader.Create(path/*, opt*/);

            while (rd.Read())
            {
                for (int l = 0; l < rd.FieldCount; l++)
                {
                    rd.GetString(l);
                }
            }
        }
    }
 
}
