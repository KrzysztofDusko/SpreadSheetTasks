using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace SpreadSheetTasks
{
    public abstract class ExcelWriter : IDisposable
    {
        internal static readonly string[] _stringDb = { "nvarchar", "varchar", "char" };
        internal static readonly Type[] _stringTypes = { typeof(String), typeof(Char), typeof(Boolean) };

        internal static readonly Type[] _numberTypes = 
        {
            typeof(sbyte), typeof(byte)
            , typeof(Int16), typeof(UInt16)
            , typeof(Int32), typeof(UInt32)
            , typeof(Int64), typeof(UInt64)
            , typeof(Single), typeof(Double)
            , typeof(Decimal)
        };

        private static readonly string[] _DbNumbers = 
        {
            "integer", "bigint"
            , "numeric", "Decimal"
            , "Double", "Single"
            , "Sbyte", "Byte"
            , "Int16", "Int32"
            , "Int64", "UInt16"
            , "UInt32", "UInt64"
        };

        public string DocPopertyProgramName { get; set; }

        internal static void SetTypes(DataColReader _dataColReader, int[] typesArray, TypeCode[] newTypes, int ColumnCount, bool detectBoolenaType = false)
        {
            if (_dataColReader.dataReader != null)
            {
                var rdr = _dataColReader.dataReader;
                for (int j = 0; j < ColumnCount; j++)
                {
                    var tempType = rdr.GetFieldType(j);
                    newTypes[j] = Type.GetTypeCode(tempType);
                    if (detectBoolenaType && tempType == typeof(Boolean))
                    {
                        typesArray[j] = 4;
                    }
                    else if (_stringTypes.Contains(tempType))
                    {
                        typesArray[j] = 0;
                    }
                    else if (_numberTypes.Contains(tempType))
                    {
                        typesArray[j] = 1;
                    }
                    else if (tempType == typeof(System.DateTime) && _dataColReader.DatabaseTypes[j].EndsWith("Date", StringComparison.OrdinalIgnoreCase))
                    {
                        typesArray[j] = 2;
                    }
                    else if (tempType == typeof(System.DateTime)
                        && (_dataColReader.DatabaseTypes[j].Equals("timestamp", StringComparison.OrdinalIgnoreCase) || _dataColReader.DatabaseTypes[j].EndsWith("DateTime", StringComparison.OrdinalIgnoreCase)))
                    {
                        typesArray[j] = 3;
                    }
                    else if (tempType == typeof(System.TimeSpan))
{
                        typesArray[j] = 3;
                    }
                    else if (tempType == typeof(Memory<byte>))
                    {
                        typesArray[j] = 5;
                    }
                    else // String, other -> as String
                    {
                        typesArray[j] = 0;
                        //throw new Exception("Excel type problem !");
                        //typesArray[j] = -1;
                    }
                }
            }
            else if (_dataColReader.DataTable != null)
            {
                var dt = _dataColReader.DataTable;
                for (int j = 0; j < ColumnCount; j++)
                {
                    newTypes[j] = Type.GetTypeCode(dt.Columns[j].DataType);       
                    if (detectBoolenaType && dt.Columns[j].DataType == typeof(Boolean))
                    {
                        typesArray[j] = 4;
                    }
                    else if (_stringTypes.Contains(dt.Columns[j].DataType))
                    {
                        typesArray[j] = 0;
                    }
                    else if (_numberTypes.Contains(dt.Columns[j].DataType))
                    {
                        typesArray[j] = 1;
                    }
                    else if (dt.Columns[j].DataType == typeof(System.DateTime))
                    {
                        typesArray[j] = 3;
                    }
                    else if (dt.Columns[j].DataType == typeof(System.TimeSpan))
{
                        typesArray[j] = 3;
                    }
                    else if (dt.Columns[j].DataType == typeof(Memory<byte>))
                    {
                        typesArray[j] = 5;
                    }
                    else // Boolean, String, other -> as String
                    {
                        typesArray[j] = 0;
                        //throw new Exception("Excel type problem !");
                        //typesArray[j] = -1;
                    }
                }
            }
            else
            {
                for (int j = 0; j < ColumnCount; j++)
                {
                    newTypes[j] = Type.GetTypeCode(_dataColReader.GetValue(j).GetType());
                    if (detectBoolenaType && _dataColReader.GetValue(j).GetType() == typeof(Boolean))
                    {
                        typesArray[j] = 4;
                    }
                    else if (_stringTypes.Contains(_dataColReader.GetValue(j).GetType()) || _stringDb.Contains(_dataColReader.DatabaseTypes[j]))
                    {
                        typesArray[j] = 0;
                    }
                    else if (_numberTypes.Contains(_dataColReader.GetValue(j).GetType()) || _DbNumbers.Contains(_dataColReader.DatabaseTypes[j]))
                    {
                        typesArray[j] = 1;
                    }
                    else if (_dataColReader.DatabaseTypes[j].Equals("Date", StringComparison.OrdinalIgnoreCase))
                    {
                        typesArray[j] = 2;
                    }
                    else if (_dataColReader.GetValue(j).GetType() == typeof(System.DateTime) || _dataColReader.DatabaseTypes[j] == "timestamp" || _dataColReader.DatabaseTypes[j] == "DateTime" /*|| kolekcjaDanych.typyBazy[j] == "TimeSpan"*/)
                    {
                        typesArray[j] = 3;
                    }
                    else if (_dataColReader.GetValue(j).GetType() == typeof(System.TimeSpan))
                    {
                        typesArray[j] = 3;
                    }
                    else if (_dataColReader.GetValue(j).GetType() == typeof(Memory<byte>))
                    {
                        typesArray[j] = 5;
                    }
                    else // other
                    {
                        throw new Exception("Excel type problem !");
                        //typesArray[j] = -1;
                    }
                }
            }
        }

        internal FileStream _newExcelFileStream;
        internal ZipArchive _excelArchiveFile;
        internal List<(string name, string pathInArchive, string pathOnDisc, bool isHidden, string nameInArchive, int sheetId)> _sheetList;

        internal string _path;
        internal const int _MAX_WIDTH = 80;
        internal int _sstCntUnique = 0;
        internal int _sstCntAll = 0;
        internal int sheetCnt = -1;
        internal DataColReader _dataColReader;
        internal bool areHeaders = false;
        internal Dictionary<string, int> _sstDic;

        internal double[] colWidesArray;
        internal int[] typesArray;
        internal TypeCode[] newTypes;

        protected int _rowsCount;
        public int RowsCount { get => _rowsCount; }

        internal abstract void FinalizeFile();
        public abstract void AddSheet(string sheetName, bool hidden = false);
        public abstract void WriteSheet(IDataReader dataReader, Boolean headers = true, int overLimit = -1, int startingRow = 0, int startingColumn = 0);

        public virtual void WriteSheet(DataTable dataTable, Boolean headers = true, int overLimit = -1, int startingRow = 0, int startingColumn = 0)
        {
            WriteSheet(dataTable.CreateDataReader(), headers, overLimit, startingRow, startingColumn);
        }

        public abstract void WriteSheet(string[] oneColumn);

        public virtual void Save()
        {
            FinalizeFile();
            _excelArchiveFile.Dispose();
            _newExcelFileStream.Dispose();
        }

        public event Action OnCompress;
        internal void DoOnCompress()
        {
            OnCompress?.Invoke();
        }
        public event Action<int> On10k;
        internal void DoOn10k(int arg)
        {
            On10k?.Invoke(arg);
        }
        public abstract void Dispose();
        public bool SuppressSomeDate { get; set; }

        internal void SetColsLengtth(int ColumnCount, object[] arr)
        {
            for (int l = 1; l <= ColumnCount; l++)
            {
                if (arr[l - 1] != null)
                {
                    var itm = arr[l - 1];
                    int lenn = 0;
                    if (itm is Memory<byte> mem)
                    {
                        lenn = mem.Length;
                    }
                    else
                    {
                        lenn = arr[l - 1].ToString().Length;
                    }
                    if (colWidesArray[l - 1] < 1.25 * lenn + 2)
                    {
                        colWidesArray[l - 1] = 1.25 * lenn + 2;
                    }
                }
            }
        }

    }

}

