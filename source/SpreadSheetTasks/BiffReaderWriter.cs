using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace SpreadSheetTasks
{
    internal sealed class BiffReaderWriter : IDisposable
    {
        //private const int WorkbookPr = 0x99;
        private const int _sheet = 0x9C; // 156

        private const int _xf = 0x2f;

        private const int _cellXfStart = 0x269;
        private const int _cellXfEnd = 0x26a;

        private const int _cellStyleXfStart = 0x272;
        private const int _cellStyleXfEnd = 0x273;

        private const int _numberFormatStart = 0x267;
        private const int _numberFormat = 0x2c;
        private const int _numberFormatEnd = 0x268;

        private const int _sharedStringStart = 159;
        private const int _stringItem = 0x13; //19

        private const uint _row = 0x00;
        private const uint _blank = 0x01;
        private const uint _number = 0x02; // BrtCellRk
        private const uint _boolError = 0x03;
        private const uint _bool = 0x04;
        private const uint _float = 0x05;
        private const uint _string = 0x06;
        private const uint _sharedString = 0x07;
        private const uint _formulaString = 0x08;
        private const uint _formulaNumber = 0x09;
        private const uint _formulaBool = 0x0a;
        private const uint _formulaError = 0x0b;

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

        internal uint _workbookId;
        internal string _recId;
        internal string _workbookName;
        internal bool _isSheet;

        internal bool ReadWorkbook()
        {
            if (!TryReadVariableValue(out var recordId) ||
                !TryReadVariableValue(out var recordLength))
                return false;
            byte[] buffer = recordLength < _buffer.Length ? _buffer : new byte[recordLength];
            if (Stream.Read(buffer, 0, (int)recordLength) != recordLength)
                return false;

            _isSheet = false;
            if (recordId == _sheet)
            {
                _workbookId = GetDWord(buffer, 4);

                uint offset = 8;
                _recId = GetNullableString(buffer, ref offset);

                // Must be between 1 and 31 characters
                uint nameLength = GetDWord(buffer, offset);
                _workbookName = GetString(buffer, offset + 4, nameLength);
                _isSheet = true;
            }
            return true;
        }

        internal bool _inCellXf;
        internal bool _inCellStyleXf;
        internal bool _inNumberFormat;

        internal ushort _parentCellStyleXf;
        internal ushort _numberFormatIndex;
        //public ushort FontIndex;

        internal int _format;
        internal string _formatString;

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
                case _cellXfStart:
                    _inCellXf = true;
                    break;
                case _cellXfEnd:
                    _inCellXf = false;
                    break;
                case _cellStyleXfStart:
                    _inCellStyleXf = true;
                    break;
                case _cellStyleXfEnd:
                    _inCellStyleXf = false;
                    break;
                case _numberFormatStart:
                    _inNumberFormat = true;
                    break;
                case _numberFormatEnd:
                    _inNumberFormat = false;
                    break;

                case _xf when _inCellStyleXf:
                    break;
                case _xf when _inCellXf:
                    {
                        _parentCellStyleXf = GetWord(buffer, 0);
                        _numberFormatIndex = GetWord(buffer, 2);
                        //var FontIndex = GetWord(buffer, 4);
                        break;
                    }

                case _numberFormat when _inNumberFormat:
                    {
                        // Must be between 1 and 255 characters
                        _format = GetWord(buffer, 0);
                        uint length = GetDWord(buffer, 2);
                        _formatString = GetString(buffer, 2 + 4, length);

                        break;
                    }
            }

            return true;

        }

        internal string? _sharedStringValue;
        internal uint _sharedStringUniqueCount = 0;
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

            if (recordId == _stringItem)
            {
                uint length = GetDWord(buffer, 1);
                _sharedStringValue = GetString(buffer, 1 + 4, length);
            }
            else if (recordId == _sharedStringStart)
            {
                _sharedStringUniqueCount = GetDWord(buffer, 4);
                _sharedStringValue = null;
            }
            else
            {
                _sharedStringValue = null;
            }

            return true;
        }

        //public object cellValue;
        internal CellType _cellType;
        internal int _intValue;
        internal double _doubleVal;
        internal bool _boolValue;
        internal string _stringValue;

        internal int _columnNum = -1;
        internal uint _xfIndex;
        //public bool isSharedStringVal = false;
        internal bool _readCell = false;
        internal int _rowIndex = -1;

        internal bool ReadWorksheet()
        {
            if (!TryReadVariableValue(out var recordId) ||
                !TryReadVariableValue(out var recordLength))
                return false;

            byte[] buffer = recordLength < _buffer.Length ? _buffer : new byte[recordLength];
            if (Stream.Read(buffer, 0, (int)recordLength) != recordLength)
                return false;

            _readCell = false;
            _columnNum = -1;
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
                case _row: // BrtRowHdr 0 = 0x0000
                    {
                        _rowIndex = GetInt32(buffer, 0);
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
                case _blank: //BrtCellBlank (1 = 0x0001)
                case _boolError:
                case _formulaError: // BrtFmlaError (11 = 0x000B)
                    //return ReadCell(null, (CellError)buffer[8]);
                    //cellValue = null;
                    _readCell = true;
                    _cellType = CellType.nullValue;
                    break;
                case _number:
                    //return ReadCell(GetRkNumber(buffer, 8));
                    //cellValue = GetRkNumber(buffer, 8);
                    _doubleVal = GetRkNumber(buffer, 8);
                    _readCell = true;
                    _cellType = CellType.doubleVal;
                    break;
                case _bool:
                case _formulaBool:
                    //return ReadCell(buffer[8] == 1);
                    //cellValue = (buffer[8] == 1);
                    _boolValue = (buffer[8] == 1);
                    _readCell = true;
                    _cellType = CellType.boolVal;
                    break;
                case _formulaNumber:
                case _float:
                    //return ReadCell(GetDouble(buffer, 8));
                    //cellValue = GetDouble(buffer, 8);
                    _doubleVal = GetDouble(buffer, 8);
                    _readCell = true;
                    _cellType = CellType.doubleVal;
                    break;
                case _string:
                case _formulaString:
                    {
                        // Must be less than 32768 characters
                        var length = GetDWord(buffer, 8);
                        //return ReadCell(GetString(buffer, 8 + 4, length));
                        //cellValue = GetString(buffer, 8 + 4, length);
                        _stringValue = GetString(buffer, 8 + 4, length);
                        _readCell = true;
                        _cellType = CellType.stringVal;
                        break;
                    }
                case _sharedString:
                    //return ReadCell((int)GetDWord(buffer, 8));
                    //cellValue = (int)GetDWord(buffer, 8);
                    _intValue = (int)GetDWord(buffer, 8);
                    _readCell = true;
                    //isSharedStringVal = true;
                    _cellType = CellType.sharedString;
                    break;
            }

            if (_readCell)
            {
                _columnNum = (int)GetDWord(buffer, 0);
                _xfIndex = GetDWord(buffer, 4) & 0xffffff;
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

        public override bool Equals(object? obj)
        {
            return obj is BiffReaderWriter writer &&
                   _workbookId == writer._workbookId;
        }
        //void Dispose(bool disposing)
        //{
        //    if (disposing)
        //        Stream.Dispose();
        //}
    }

    internal class DataColReader
    {
        internal readonly IDataReader _dataReader;
        internal DataTable _dataTable;
        private readonly object[,] _tabelarData;
        private readonly bool _isDataReader;
        private readonly bool _isDataTable;
        internal int _dataTableRowsCount;

        private readonly bool _headers;
        private int _rowNum = 0;

        internal string[] _databaseTypes;

        public DataColReader(IDataReader reader, Boolean headers = false, int overLimit = -1)
        {
            this._dataReader = reader;
            this._headers = headers;
            this._isDataReader = true;
            this._overLimit = overLimit;

            _databaseTypes = new string[_dataReader.FieldCount];
            for (int i = 0; i < _databaseTypes.Length; i++)
            {
                _databaseTypes[i] = _dataReader.GetDataTypeName(i);
            }
        }

        public DataColReader(DataTable dataTable, Boolean headers = false, int overLimit = -1)
        {
            this._dataTable = dataTable;
            this._headers = headers;
            this._isDataTable = true;
            this._overLimit = overLimit;
            this._dataTableRowsCount = dataTable.Rows.Count;

            _databaseTypes = new string[_dataTable.Columns.Count];

            // WORK TO DO !!
            for (int i = 0; i < _databaseTypes.Length; i++)
            {
                _databaseTypes[i] = _dataTable.Columns[i].DataType.ToString();
            }
        }

        public DataColReader(string[,] tabelarData)
        {
            this._tabelarData = tabelarData;
            _isDataReader = false;
            _databaseTypes = new string[tabelarData.Length];
            for (int i = 0; i < _databaseTypes.Length; i++)
            {
                _databaseTypes[i] = "-1";
            }
        }

        private readonly int _overLimit = -1;
        public int FieldCount    // the Name property
        {
            get
            {
                if (_isDataReader && _overLimit > 0)
                {
                    return _overLimit;
                }
                else if (_isDataReader)
                {
                    return _dataReader.FieldCount;
                }
                else if (_isDataTable)
                {
                    return _dataTable.Columns.Count;
                }
                else
                {
                    return _tabelarData.GetUpperBound(1) + 1;
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
                else if (top100 != null && _topNum <= top100.Count)
                {
                    _topNum++;
                    if (_topNum == top100.Count + 1)
                    {
                        top100 = null;
                        return AreNextRows;
                    }
                    return true;
                }
                else
                {
                    return _dataReader.Read();
                }
            }
            else if (_isDataTable)
            {
                if (_rowNum >= 2 && _rowNum < _dataTableRowsCount + 2)
                {
                    _dataTableRow = _dataTable.Rows[_rowNum - 2].ItemArray;
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
                return (_rowNum < _tabelarData.GetUpperBound(0) + 2);
            }
        }

        private object[]? _dataTableRow;
        public object GetValue(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetValue(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return _dataReader.GetName(j);
                }
                else
                {
                    return top100[_topNum - 1][j];
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
                    return _dataTable.Columns[j].ColumnName;
                }
            }
            else
            {
                return _tabelarData[_rowNum - 1, j];
            }
        }

        
        public bool GetBoolean(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetBoolean(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("bool for header ?");
                }
                else
                {
                    return (bool)top100[_topNum - 1][j];
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
                return (bool)_tabelarData[_rowNum - 1, j];
            }
        }

        public char GetChar(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetChar(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("char for header ?");
                }
                else
                {
                    return (char)top100[_topNum - 1][j];
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
                return (char)_tabelarData[_rowNum - 1, j];
            }
        }

        public byte GetByte(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetByte(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("byte for header ?");
                }
                else
                {
                    return (byte)top100[_topNum - 1][j];
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
                return (byte)_tabelarData[_rowNum - 1, j];
            }
        }

        public sbyte GetSByte(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return (sbyte)_dataReader.GetValue(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("byte for header ?");
                }
                else
                {
                    return (sbyte)top100[_topNum - 1][j];
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
                return (sbyte)_tabelarData[_rowNum - 1, j];
            }
        }

        public Int16 GetInt16(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetInt16(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int16 for header ?");
                }
                else
                {
                    return (Int16)top100[_topNum - 1][j];
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
                return (Int16)_tabelarData[_rowNum - 1, j];
            }
        }

        public Int32 GetInt32(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetInt32(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int32 for header ?");
                }
                else
                {
                    return (Int32)top100[_topNum - 1][j];
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
                return (Int32)_tabelarData[_rowNum - 1, j];
            }
        }

        public Int64 GetInt64(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetInt64(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("Int64 for header ?");
                }
                else
                {
                    return (Int64)top100[_topNum - 1][j];
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
                return (Int32)_tabelarData[_rowNum - 1, j];
            }
        }

        public float GetFloat(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetFloat(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("float for header ?");
                }
                else
                {
                    return (float)top100[_topNum - 1][j];
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
                return (float)_tabelarData[_rowNum - 1, j];
            }
        }
        public double GetDouble(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetDouble(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("double for header ?");
                }
                else
                {
                    return (double)top100[_topNum - 1][j];
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
                return (double)_tabelarData[_rowNum - 1, j];
            }
        }
        public decimal GetDecimal(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetDecimal(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("decimal for header ?");
                }
                else
                {
                    return (decimal)top100[_topNum - 1][j];
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
                return (decimal)_tabelarData[_rowNum - 1, j];
            }
        }

        public DateTime GetDateTime(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetDateTime(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    throw new Exception("DateTime for header ?");
                }
                else
                {
                    return (DateTime)top100[_topNum - 1][j];
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
                return (DateTime)_tabelarData[_rowNum - 1, j];
            }
        }

        public string GetString(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.GetString(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return _dataReader.GetName(j);
                }
                else
                {
                    return top100[_topNum - 1][j].ToString();
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
                    return _dataTable.Columns[j].ColumnName;
                }
            }
            else
            {
                return _tabelarData[_rowNum - 1, j].ToString();
            }
        }

        public bool IsDBNull(int j)
        {
            if (_isDataReader)
            {
                if ((_rowNum > 1 || !_headers) && top100 == null)
                {
                    return _dataReader.IsDBNull(j);
                }
                else if (_headers && _rowNum == 1)
                {
                    return false;
                }
                else
                {
                    return top100[_topNum - 1][j] == null || top100[_topNum - 1][j] == DBNull.Value;
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
                    return _dataTable.Columns[j].ColumnName == null;
                }
            }
            else
            {
                return _tabelarData[_rowNum - 1, j] == null || _tabelarData[_rowNum - 1, j] == DBNull.Value;
            }
        }

        public void GetWidthFromDataTable(Span<double> width, double maxWidth, bool doAutofilter)
        {
            int n = _dataTableRowsCount > 100 ? 100 : _dataTableRowsCount;
            int m = FieldCount;

            for (int j = 0; j < m; j++)
            {
                double valTemp = 1.25 * _dataTable.Columns[j].ToString().Length + 2;
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
                    double valTemp = 1.25 * _dataTable.Rows[i][j].ToString().Length + 2;
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
        private int _topNum = 0;
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
