using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace SpreadSheetTasks
{
    internal sealed class BiffReaderWriter : IDisposable
    {
        //private const int WorkbookPr = 0x99;
        private const int Sheet = 0x9C; // 156

        private const int Xf = 0x2f;

        private const int CellXfStart = 0x269;
        private const int CellXfEnd = 0x26a;

        private const int CellStyleXfStart = 0x272;
        private const int CellStyleXfEnd = 0x273;

        private const int NumberFormatStart = 0x267;
        private const int NumberFormat = 0x2c;
        private const int NumberFormatEnd = 0x268;

        private const int SharedStringStart = 159;
        private const int StringItem = 0x13; //19

        private const uint Row = 0x00;
        private const uint Blank = 0x01;
        private const uint Number = 0x02; // BrtCellRk
        private const uint BoolError = 0x03;
        private const uint Bool = 0x04;
        private const uint Float = 0x05;
        private const uint String = 0x06;
        private const uint SharedString = 0x07;
        private const uint FormulaString = 0x08;
        private const uint FormulaNumber = 0x09;
        private const uint FormulaBool = 0x0a;
        private const uint FormulaError = 0x0b;

        // private const uint WorksheetBegin = 0x81;
        // private const uint WorksheetEnd = 0x82;
        //private const uint SheetDataBegin = 0x91;
        //private const uint SheetDataEnd = 0x92;
        //private const uint SheetPr = 0x93; // == BrtWsProp
        //private const uint SheetFormatPr = 0x1E5;

        // private const uint ColumnsBegin = 0x186;
        //private const uint Column = 0x3C; // column info

        // private const uint ColumnsEnd = 0x187;
        //private const uint HeaderFooter = 0x1DF;

        // private const uint MergeCellsBegin = 0x00B1; //177
        // private const uint MergeCellsEnd = 0x00B2; //178
        //private const uint MergeCell = 0x00B0; // 176

        //private const uint BrtBeginSheet = 0x0081; // 129
        //private const uint BrtWsProp = 0x0093; // 147 // SheetPr
        //private const uint LHRecord = 0x0094; // 148
        //private const uint BrtBeginWsViews = 0x0085;//133
        //private const uint BrtBeginWsView = 0x0089; // 137
        //private const uint BrtSel = 0x0098; // 152
        //private const uint BrtEndWsView = 0x008A; // 138
        //private const uint BrtEndWsViews = 0x0086; //134

        //private const uint BrtACBegin = 0x0025;// 37
        //private const uint BrtWsFmtInfoEx14 = 0x0415;//1045
        //private const uint BrtACEnd = 0x0026;//38
        //private const uint BrtWsFmtInfo = 0x01E5;//485

        //private const uint BrtBeginSheetData = 0x0091;//145
        //private const uint BrtRwDescent = 0x0400;//1024
        //private const uint BrtEndSheetData = 0x0092;//146

        //private const uint BrtSheetProtection = 0x0217;//535
        //private const uint BrtPhoneticInfo = 0x0219;//537
        //private const uint BrtPrintOptions = 0x01DD;//477
        //private const uint BrtMargins = 0x01DC;//476
        //private const uint BrtUid = 0x0C00;//3072
        //private const uint BrtEndSheet = 0x0082;//130


        private readonly byte[] _buffer = new byte[128];
        Stream Stream { get; }

        public BiffReaderWriter(Stream stream)
        {
            Stream = stream ?? throw new ArgumentNullException(nameof(stream));
        }

        private enum SheetVisibility : byte
        {
            Visible = 0x0,
            Hidden = 0x1,
            VeryHidden = 0x2
        }

        internal uint workbookId;
        internal string recId;
        internal string workbookName;
        internal bool isSheet;

        internal bool ReadWorkbook()
        {
            if (!TryReadVariableValue(out var recordId) ||
                !TryReadVariableValue(out var recordLength))
                return false;
            byte[] buffer = recordLength < _buffer.Length ? _buffer : new byte[recordLength];
            if (Stream.Read(buffer, 0, (int)recordLength) != recordLength)
                return false;

            isSheet = false;
            if (recordId == Sheet)
            {
                workbookId = GetDWord(buffer, 4);

                uint offset = 8;
                recId = GetNullableString(buffer, ref offset);

                // Must be between 1 and 31 characters
                uint nameLength = GetDWord(buffer, offset);
                workbookName = GetString(buffer, offset + 4, nameLength);
                isSheet = true;
            }
            return true;
        }

        internal bool _inCellXf;
        internal bool _inCellStyleXf;
        internal bool _inNumberFormat;

        internal ushort ParentCellStyleXf;
        internal ushort NumberFormatIndex;
        //public ushort FontIndex;

        internal int format;
        internal string formatString;

        public bool ReadStyles()
        {
            if (!TryReadVariableValue(out var recordId) ||
                !TryReadVariableValue(out var recordLength))
                return false;

            byte[] buffer = recordLength < _buffer.Length ? _buffer : new byte[recordLength];
            if (Stream.Read(buffer, 0, (int)recordLength) != recordLength)
                return false;


            switch (recordId)
            {
                case CellXfStart:
                    _inCellXf = true;
                    break;
                case CellXfEnd:
                    _inCellXf = false;
                    break;
                case CellStyleXfStart:
                    _inCellStyleXf = true;
                    break;
                case CellStyleXfEnd:
                    _inCellStyleXf = false;
                    break;
                case NumberFormatStart:
                    _inNumberFormat = true;
                    break;
                case NumberFormatEnd:
                    _inNumberFormat = false;
                    break;

                case Xf when _inCellStyleXf:
                    break;
                case Xf when _inCellXf:
                    {
                        ParentCellStyleXf = GetWord(buffer, 0);
                        NumberFormatIndex = GetWord(buffer, 2);
                        //var FontIndex = GetWord(buffer, 4);
                        break;
                    }

                case NumberFormat when _inNumberFormat:
                    {
                        // Must be between 1 and 255 characters
                        format = GetWord(buffer, 0);
                        uint length = GetDWord(buffer, 2);
                        formatString = GetString(buffer, 2 + 4, length);

                        break;
                    }
            }

            return true;

        }

        internal string sharedStringValue;
        internal uint sharedStringUniqueCount = 0;
        public bool ReadSharedStrings()
        {
            if (!TryReadVariableValue(out var recordId) ||
                !TryReadVariableValue(out var recordLength))
                return false;

            byte[] buffer = recordLength < _buffer.Length ? _buffer : new byte[recordLength];

            //if (Stream.Read(buffer, 0, (int)recordLength) != recordLength)
            //    return false;

            uint readed = 0;
            do
            {
                readed += (uint)Stream.Read(buffer, (int)readed, (int)(recordLength - readed));
                if (readed == 0)
                {
                    return false;
                }
            } while (readed < recordLength);

            if (recordId == StringItem)
            {
                uint length = GetDWord(buffer, 1);
                sharedStringValue = GetString(buffer, 1 + 4, length);
            }
            else if (recordId == SharedStringStart)
            {
                sharedStringUniqueCount = GetDWord(buffer, 4);
                sharedStringValue = null;
            }
            else
            {
                sharedStringValue = null;
            }

            return true;
        }

        //public object cellValue;
        internal CellType cellType;
        internal int intValue;
        internal double doubleVal;
        internal bool boolValue;
        internal string stringValue;

        internal int columnNum = -1;
        internal uint xfIndex;
        //public bool isSharedStringVal = false;
        internal bool readCell = false;
        internal int rowIndex = -1;

        internal bool ReadWorksheet()
        {
            if (!TryReadVariableValue(out var recordId) ||
                !TryReadVariableValue(out var recordLength))
                return false;

            byte[] buffer = recordLength < _buffer.Length ? _buffer : new byte[recordLength];
            if (Stream.Read(buffer, 0, (int)recordLength) != recordLength)
                return false;

            readCell = false;
            columnNum = -1;
            //isSharedStringVal = false;

            switch (recordId)
            {
                //case BrtEndWsViews:
                //    break;
                //case BrtSel:
                //    break;
                //case SheetDataBegin:
                //sheetDataBeginRecord = true;
                //break;
                //case SheetDataEnd:
                //sheetDataBeginRecord = false;
                //sheetDataEndRecord = true;
                //break;
                //case SheetPr: // BrtWsProp
                //    {
                //        // Must be between 0 and 31 characters
                //        uint length = GetDWord(buffer, 19);

                //        // To behave the same as when reading an xml based file. 
                //        // GetAttribute returns null both if the attribute is missing
                //        // or if it is empty.
                //        string codeName = length == 0 ? null : GetString(buffer, 19 + 4, length);
                //        //return new SheetPrRecord(codeName);
                //        break;
                //    }
                //break;
                //case SheetFormatPr: // BrtWsFmtInfo 
                //{
                //    // TODO Default column width
                //    var unsynced = (buffer[8] & 0b1000) != 0;
                //    uint? defaultHeight = null;
                //    if (unsynced)
                //        defaultHeight = GetWord(buffer, 6);
                //    //return new SheetFormatPrRecord(defaultHeight);
                //    break;
                //}
                //break;
                //case Column: // BrtColInfo 
                //    {
                //        int minimum = GetInt32(buffer, 0);
                //        int maximum = GetInt32(buffer, 4);
                //        byte flags = buffer[16];
                //        bool hidden = (flags & 0b1) != 0;
                //        bool unsynced = (flags & 0b10) != 0;

                //        double? width = null;
                //        if (unsynced)
                //            width = GetDWord(buffer, 8) / 256.0;
                //        //return new ColumnRecord(new Column(minimum, maximum, hidden, width));
                //        break;
                //        //{0,0,0,0,0,0,0,36,59,0,0,0,0,0,0,2}
                //    }
                //break;
                //case HeaderFooter: // BrtBeginHeaderFooter 
                //{
                //    var flags = buffer[0];
                //    bool differentOddEven = (flags & 1) != 0;
                //    bool differentFirst = (flags & 0b10) != 0;
                //    uint offset = 2;
                //    var header = GetNullableString(buffer, ref offset);
                //    var footer = GetNullableString(buffer, ref offset);
                //    var headerEven = GetNullableString(buffer, ref offset);
                //    var footerEven = GetNullableString(buffer, ref offset);
                //    var headerFirst = GetNullableString(buffer, ref offset);
                //    var footerFirst = GetNullableString(buffer, ref offset);
                //    break;
                //}
                //break;
                //case BrtBeginSheetData:
                //    Console.WriteLine("posiotion of BrtBeginSheetData");
                //    Console.WriteLine(Stream.Position);
                //    break;
                //case BrtEndSheetData:
                //    Console.WriteLine("posiotion of BrtEndSheetData");
                //    Console.WriteLine(Stream.Position);
                //    break;
                //case BrtACBegin:
                //    Console.WriteLine("posiotion of BrtACBegin");
                //    Console.WriteLine(Stream.Position);
                //    break;
                //case BrtACEnd:
                //    Console.WriteLine("posiotion of BrtACEnd");
                //    Console.WriteLine(Stream.Position);
                //    break;

                //case BrtRwDescent:
                //    Console.WriteLine("posiotion of BrtRwDescent");
                //    Console.WriteLine(Stream.Position);
                //    break;
                //case MergeCell:
                //int fromRow = GetInt32(buffer, 0);
                //int toRow = GetInt32(buffer, 4);
                //int fromColumn = GetInt32(buffer, 8);
                //int toColumn = GetInt32(buffer, 12);
                //break;
                case Row: // BrtRowHdr 0 = 0x0000
                    {
                        rowIndex = GetInt32(buffer, 0);
                        //    byte flags = buffer[11];
                        //    bool hidden = (flags & 0b10000) != 0;
                        //    bool unsynced = (flags & 0b100000) != 0;

                        //    double? height = null;
                        //    if (unsynced)
                        //        height = GetWord(buffer, 8) / 20.0; // Where does 20 come from?

                        //    // TODO: Default format ?
                        break;
                    }
                //case Blank: //BrtCellBlank
                //return ReadCell(null);
                //cellValue = null; 
                //readCell = true;
                //break;
                case Blank: //BrtCellBlank (1 = 0x0001)
                case BoolError:
                case FormulaError: // BrtFmlaError (11 = 0x000B)
                    //return ReadCell(null, (CellError)buffer[8]);
                    //cellValue = null;
                    readCell = true;
                    cellType = CellType.nullValue;
                    break;
                case Number:
                    //return ReadCell(GetRkNumber(buffer, 8));
                    //cellValue = GetRkNumber(buffer, 8);
                    doubleVal = GetRkNumber(buffer, 8);
                    readCell = true;
                    cellType = CellType.doubleVal;
                    break;
                case Bool:
                case FormulaBool:
                    //return ReadCell(buffer[8] == 1);
                    //cellValue = (buffer[8] == 1);
                    boolValue = (buffer[8] == 1);
                    readCell = true;
                    cellType = CellType.boolVal;
                    break;
                case FormulaNumber:
                case Float:
                    //return ReadCell(GetDouble(buffer, 8));
                    //cellValue = GetDouble(buffer, 8);
                    doubleVal = GetDouble(buffer, 8);
                    readCell = true;
                    cellType = CellType.doubleVal;
                    break;
                case String:
                case FormulaString:
                    {
                        // Must be less than 32768 characters
                        var length = GetDWord(buffer, 8);
                        //return ReadCell(GetString(buffer, 8 + 4, length));
                        //cellValue = GetString(buffer, 8 + 4, length);
                        stringValue = GetString(buffer, 8 + 4, length);
                        readCell = true;
                        cellType = CellType.stringVal;
                        break;
                    }
                case SharedString:
                    //return ReadCell((int)GetDWord(buffer, 8));
                    //cellValue = (int)GetDWord(buffer, 8);
                    intValue = (int)GetDWord(buffer, 8);
                    readCell = true;
                    //isSharedStringVal = true;
                    cellType = CellType.sharedString;
                    break;
            }

            if (readCell)
            {
                columnNum = (int)GetDWord(buffer, 0);
                xfIndex = GetDWord(buffer, 4) & 0xffffff;
            }

            return true;
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static uint GetDWord(byte[] buffer, uint offset)
        {
            uint result = (uint)buffer[offset + 3] << 24;
            result += (uint)buffer[offset + 2] << 16;
            result += (uint)buffer[offset + 1] << 8;
            result += buffer[offset];
            return result;
        }


        //https://github.com/ExcelDataReader/ExcelDataReader
        static int GetInt32(byte[] buffer, uint offset)
        {
            int result = buffer[offset + 3] << 24;
            result += buffer[offset + 2] << 16;
            result += buffer[offset + 1] << 8;
            result += buffer[offset];
            return result;
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static ushort GetWord(byte[] buffer, uint offset)
        {
            ushort result = (ushort)(buffer[offset + 1] << 8);
            result += buffer[offset];
            return result;
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        /*public static string GetString(byte[] buffer, uint offset, uint length)
        {
            StringBuilder sb = new StringBuilder((int)length);
            for (uint i = offset; i < offset + 2 * length; i += 2)
                sb.Append((char)GetWord(buffer, i));
            return sb.ToString();
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static string? GetNullableString(byte[] buffer, ref uint offset)
        {
            var length = GetDWord(buffer, offset);
            offset += 4;
            if (length == uint.MaxValue)
                return null;
            StringBuilder sb = new StringBuilder((int)length);
            uint end = offset + length * 2;
            for (; offset < end; offset += 2)
                sb.Append((char)GetWord(buffer, offset));
            return sb.ToString();
        }*/


        private static string GetString(byte[] buffer, uint offset, uint length)
        {
            //https://docs.microsoft.com/en-US/dotnet/api/system.string.create?view=net-5.0
            return string.Create((int)length, (buffer, offset, length), (chars, state) =>
            {
                int l = 0;
                byte[] buff = state.buffer;
                for (uint i = state.offset; i < state.offset + 2 * state.length; i += 2)
                    chars[l++] = (char)GetWord(buff, i);
            });


            //Span<char> array = stackalloc char[(int)length];
            //int l = 0;
            //for (uint i = offset; i < offset + 2 * length; i += 2)
            //    array[l++] = (char)GetWord(buffer, i);

            //return new string(array);

            //char[] array = ArrayPool<char>.Shared.Rent((int)length);
            //int l = 0;
            //for (uint i = offset; i < offset + 2 * length; i += 2)
            //    array[l++] = (char)GetWord(buffer, i);
            //string s1 = new string(array.AsSpan().Slice(0, (int)length));
            //ArrayPool<char>.Shared.Return(array);
            //return s1;

        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static string GetNullableString(byte[] buffer, ref uint offset)
        {
            var length = GetDWord(buffer, offset);
            offset += 4;
            if (length == uint.MaxValue)
                return null;

            char[] array = new char[length];
            int l = 0;

            uint end = offset + length * 2;
            for (; offset < end; offset += 2)
                array[l++] = (char)GetWord(buffer, offset);
            return new string(array);
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        //2.5.122 RkNumber
        static double GetRkNumber(byte[] buffer, uint offset)
        {
            double result;

            byte flags = buffer[offset];

            if ((flags & 0x02) != 0)
            {
                result = GetInt32(buffer, offset) >> 2;
            }
            else
            {
                result = BitConverter.Int64BitsToDouble((GetDWord(buffer, offset) & -4) << 32);
            }

            if ((flags & 0x01) != 0)
            {
                result /= 100;
            }

            return result;
        }

        //https://github.com/ExcelDataReader/ExcelDataReader
        static double GetDouble(byte[] buffer, uint offset)
        {
            uint num = GetDWord(buffer, offset);
            uint num2 = GetDWord(buffer, offset + 4);
            long num3 = ((long)num2 << 32) | num;
            return BitConverter.Int64BitsToDouble(num3);
        }

        //https://github.com/ExcelDataReader/ExcelDataReader

        private bool TryReadVariableValue(out uint value)
        {
            value = 0;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;

            byte b1 = _buffer[0];
            value = (uint)(b1 & 0x7F);

            if ((b1 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b2 = _buffer[0];
            value = ((uint)(b2 & 0x7F) << 7) | value;

            if ((b2 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b3 = _buffer[0];
            value = ((uint)(b3 & 0x7F) << 14) | value;

            if ((b3 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b4 = _buffer[0];
            value = ((uint)(b4 & 0x7F) << 21) | value;

            return true;
        }

        public void Dispose()
        {
            Stream.Dispose();
        }
        //void Dispose(bool disposing)
        //{
        //    if (disposing)
        //        Stream.Dispose();
        //}
    }

    internal class DataColReader
    {
        internal IDataReader dataReader;
        internal DataTable DataTable;
        private readonly object[,] daneTabelaryczne;
        private readonly bool _isDataReader;
        private readonly bool _isDataTable;
        internal int DataTableRowsCount;

        private readonly bool _headers;
        private int _rowNum = 0;

        internal string[] DatabaseTypes;

        public DataColReader(IDataReader reader, Boolean headers = false, int overLimit = -1)
        {
            this.dataReader = reader;
            this._headers = headers;
            this._isDataReader = true;
            this.overLimit = overLimit;

            DatabaseTypes = new string[dataReader.FieldCount];
            for (int i = 0; i < DatabaseTypes.Length; i++)
            {
                DatabaseTypes[i] = dataReader.GetDataTypeName(i);
            }
        }

        public DataColReader(DataTable dataTable, Boolean headers = false, int overLimit = -1)
        {
            this.DataTable = dataTable;
            this._headers = headers;
            this._isDataTable = true;
            this.overLimit = overLimit;
            this.DataTableRowsCount = dataTable.Rows.Count;

            DatabaseTypes = new string[DataTable.Columns.Count];

            // WORK TO DO !!
            for (int i = 0; i < DatabaseTypes.Length; i++)
            {
                DatabaseTypes[i] = DataTable.Columns[i].DataType.ToString();
            }
        }

        public DataColReader(string[,] daneZtabeli)
        {
            this.daneTabelaryczne = daneZtabeli;
            _isDataReader = false;
            DatabaseTypes = new string[daneZtabeli.Length];
            for (int i = 0; i < DatabaseTypes.Length; i++)
            {
                DatabaseTypes[i] = "-1";
            }
        }

        private readonly int overLimit = -1;
        public int FieldCount    // the Name property
        {
            get
            {
                if (_isDataReader && overLimit > 0)
                {
                    return overLimit;
                }
                else if (_isDataReader)
                {
                    return dataReader.FieldCount;
                }
                else if (_isDataTable)
                {
                    return DataTable.Columns.Count;
                }
                else
                {
                    return daneTabelaryczne.GetUpperBound(1) + 1;
                }
            }
        }
        public bool Read()
        {
            ++_rowNum;

            if (_isDataReader)
            {
                if (_isDataReader && _rowNum <= 1 && _headers)
                {
                    return true;
                }
                else if (top100 != null && topNum <= top100.Count)
                {
                    topNum++;
                    if (topNum == top100.Count + 1)
                    {
                        top100 = null;
                        return AreNextRows;
                    }
                    return true;
                }
                else
                {
                    return dataReader.Read();
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum >= 2 && _rowNum < DataTableRowsCount + 2)
                {
                    _dataTableRow = DataTable.Rows[_rowNum - 2].ItemArray;
                    return true;
                }
                else if (_rowNum == 1)
                {
                    return true;
                }
                return false;
            }
            else
            {
                return (_rowNum < daneTabelaryczne.GetUpperBound(0) + 2);
            }
        }

        private object[] _dataTableRow;
        public object GetValue(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetValue(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return dataReader.GetName(j);
                }
                else
                {
                    return top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    return _dataTableRow[j];
                }
                else
                {
                    return DataTable.Columns[j].ColumnName;
                }
            }
            else
            {
                return daneTabelaryczne[_rowNum - 1, j];
            }
        }

        
        public bool GetBoolean(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetBoolean(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("bool for header ?");
                }
                else
                {
                    return (bool)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (bool)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("bool for header ?");
                }
            }
            else
            {
                return (bool)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public char GetChar(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetChar(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("char for header ?");
                }
                else
                {
                    return (char)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (char)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("char for header ?");
                }
            }
            else
            {
                return (char)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public byte GetByte(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetByte(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("byte for header ?");
                }
                else
                {
                    return (byte)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (byte)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("byte for header ?");
                }
            }
            else
            {
                return (byte)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public sbyte GetSByte(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return (sbyte)dataReader.GetValue(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("byte for header ?");
                }
                else
                {
                    return (sbyte)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (sbyte)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("sbyte for header ?");
                }
            }
            else
            {
                return (sbyte)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public Int16 GetInt16(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetInt16(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int16 for header ?");
                }
                else
                {
                    return (Int16)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (Int16)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("Int16 for header ?");
                }
            }
            else
            {
                return (Int16)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public Int32 GetInt32(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetInt32(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int32 for header ?");
                }
                else
                {
                    return (Int32)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (Int32)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("Int32 for header ?");
                }
            }
            else
            {
                return (Int32)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public Int64 GetInt64(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetInt64(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int64 for header ?");
                }
                else
                {
                    return (Int64)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (Int64)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("Int64 for header ?");
                }
            }
            else
            {
                return (Int32)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public float GetFloat(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetFloat(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("float for header ?");
                }
                else
                {
                    return (float)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (float)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("float for header ?");
                }
            }
            else
            {
                return (float)daneTabelaryczne[_rowNum - 1, j];
            }
        }
        public double GetDouble(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetDouble(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("double for header ?");
                }
                else
                {
                    return (double)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (double)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("double for header ?");
                }
            }
            else
            {
                return (double)daneTabelaryczne[_rowNum - 1, j];
            }
        }
        public decimal GetDecimal(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetDecimal(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("decimal for header ?");
                }
                else
                {
                    return (decimal)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return (decimal)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("decimal for header ?");
                }
            }
            else
            {
                return (decimal)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public DateTime GetDateTime(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetDateTime(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("DateTime for header ?");
                }
                else
                {
                    return (DateTime)top100[topNum - 1][j];
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    return (DateTime)_dataTableRow[j];
                }
                else
                {
                    throw new Exception("decimal for header ?");
                }
            }
            else
            {
                return (DateTime)daneTabelaryczne[_rowNum - 1, j];
            }
        }

        public string GetString(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.GetString(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return dataReader.GetName(j);
                }
                else
                {
                    return top100[topNum - 1][j].ToString();
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return _dataTableRow[j].ToString();
                }
                else
                {
                    return DataTable.Columns[j].ColumnName;
                }
            }
            else
            {
                return daneTabelaryczne[_rowNum - 1, j].ToString();
            }
        }

        public bool IsDBNull(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return dataReader.IsDBNull(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return false;
                }
                else
                {
                    return top100[topNum - 1][j] == null || top100[topNum - 1][j] == DBNull.Value;
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum > 1 || !_headers)
                {
                    //return DataTable.Rows[_rowNum-2][j];
                    return _dataTableRow[j] == null || _dataTableRow[j] == DBNull.Value;
                }
                else
                {
                    return DataTable.Columns[j].ColumnName == null;
                }
            }
            else
            {
                return daneTabelaryczne[_rowNum - 1, j] == null || daneTabelaryczne[_rowNum - 1, j] == DBNull.Value;
            }
        }

        public void GetWidthFromDataTable(Span<double> width, double maxWidth, bool doAutofilter)
        {
            int n = DataTableRowsCount > 100 ? 100 : DataTableRowsCount;
            int m = FieldCount;

            for (int j = 0; j < m; j++)
            {
                double valTemp = 1.25 * DataTable.Columns[j].ToString().Length + 2;
                if (doAutofilter)
                {
                    valTemp += 2;
                }

                if (valTemp > maxWidth)
                {
                    valTemp = maxWidth;
                }

                if (width[j] < valTemp)
                {
                    width[j] = valTemp;
                }
            }

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    double valTemp = 1.25 * DataTable.Rows[i][j].ToString().Length + 2;
                    if (valTemp > maxWidth)
                    {
                        valTemp = maxWidth;
                    }

                    if (width[j] < valTemp)
                    {
                        width[j] = valTemp;
                    }
                }
            }
        }
        public bool AreNextRows { get; set; }
        int topNum = 0;
        public List<object[]> top100;
    }

    //https://github.com/ExcelDataReader/ExcelDataReader
    //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/aa9f2bac-991a-42a8-8cfa-507de84017b6


    internal enum CellType
    {
        doubleVal,
        boolVal,
        stringVal,
        sharedString,
        nullValue
    }
}
