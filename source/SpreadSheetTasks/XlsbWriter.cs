using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace SpreadSheetTasks
{
    public sealed class XlsbWriter : ExcelWriter, IDisposable
    {
        private readonly static byte[] _generalStyle = [0, 0, 0, 0]; //number of style
        private readonly static byte[] _autoFilterStartBytes = [0xA1, 0x01, 0x10];
        private readonly static byte[] _autoFilterEndBytes = [0xA2, 0x01, 0x00];
        private readonly static byte[] _sheet1Bytes =
        [
            //sheet1Bytes[0..84]
            0x81,0x01,0x00,0x93,0x01,0x17,0xCB,0x04, //0 ..7
            0x02,0x00,0x40,0x00,0x00,0x00,0x00,0x00, //8 ..15
            0x00,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF, //16 ..23
            0xFF,0x00,0x00,0x00,0x00,0x94,0x01,0x10, //24 ..31
            0x00,0x00,0x00,0x00,//start row //32 (not requied)
            0x00,0x00,0x00,0x00,//last row  //36 (not requied)
            0x00,0x00,0x00,0x00,//start col //40 (not requied)
            0x00,0x00,0x00,0x00,//last col  //44 - 47 (not requied)
            0x85,0x01,0x00,0x89,0x01,0x1E,0xDC,0x03,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x40,0x00,0x00,0x00,
            0x64,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00, 

             //sheet1Bytes[84..]
                                0x98,0x01,0x24,0x03,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x8A,0x01,0x00,0x86,0x01,
            0x00,0x25,0x06,0x01,0x00,0x02,0x0E,0x00,
            0x80,0x95,0x08,0x02,0x05,0x00,0x26,0x00,
            0xE5,0x03,0x0C,0xFF,0xFF,0xFF,0xFF,0x08,
            0x00,0x2C,0x01,0x00,0x00,0x00,0x00,0x91,
            0x01,0x00,0x25,0x06,0x01,0x00,0x02,0x0E,
            0x00,0x80,0x80,0x08,0x02,0x05,0x00,0x26,
            0x00,0x00,0x19,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x2C,0x01,0x00,0x00,0x00,
            0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x07,0x0C,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x92,0x01,0x00,0x97,0x04,0x42,
            0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,
            0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,
            0x00,0x00,
            //autofilter goes here
            /*sheet1Bytes[290] = */
            0xDD,0x03,0x02,0x10,0x00,0xDC,0x03,0x30,
            0x66,0x66,0x66,0x66,0x66,0x66,0xE6,0x3F,
            0x66,0x66,0x66,0x66,0x66,0x66,0xE6,0x3F,
            0x00,0x00,0x00,0x00,0x00,0x00,0xE8,0x3F,
            0x00,0x00,0x00,0x00,0x00,0x00,0xE8,0x3F,
            0x33,0x33,0x33,0x33,0x33,0x33,0xD3,0x3F,
            0x33,0x33,0x33,0x33,0x33,0x33,0xD3,0x3F,
            0x25,0x06,0x01,0x00,0x00,0x10,0x00,0x80,
            0x80,0x18,0x10,0x00,0x00,0x00,0x00,0x01,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x26,0x00,0x82,0x01,0x00
        ];

        private readonly static byte[] _stickHeaderA1bytes =
        [
            0x97,0x01,0x1D,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0xF0,0x3F,0x01,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x02,0x00,0x00,0x00,0x03,
        ];

        private readonly static byte[] _stylesBin =
        [
            0x96,0x02,0x00,0xE7,0x04,0x04,0x02,
            0x00,0x00,0x00,0x2C,0x2C,0xA4,0x00,0x13,
            0x00,0x00,0x00,0x79,0x00,0x79,0x00,0x79,
            0x00,0x79,0x00,0x5C,0x00,0x2D,0x00,0x6D,
            0x00,0x6D,0x00,0x5C,0x00,0x2D,0x00,0x64,
            0x00,0x64,0x00,0x5C,0x00,0x20,0x00,0x68,
            0x00,0x68,0x00,0x3A,0x00,0x6D,0x00,0x6D,
            0x00,0x2C,0x1E,0xA6,0x00,0x0C,0x00,0x00,
            0x00,0x79,0x00,0x79,0x00,0x79,0x00,0x79,
            0x00,0x5C,0x00,0x2D,0x00,0x6D,0x00,0x6D,
            0x00,0x5C,0x00,0x2D,0x00,0x64,0x00,0x64,
            0x00,0xE8,0x04,0x00,0xE3,0x04,0x04,0x01,
            0x00,0x00,0x00,

            //standard font ?
            0x2B,0x27,0xDC,0x00,0x00,0x00,0x90,0x01,
            0x00,0x00,0x00,0x02,0x00,0x00,0x07,0x01,
            0x00,0x00,0x00,0x00,0x00,0xFF,0x02,0x07,
            0x00,0x00,0x00,0x43,0x00,0x61,0x00,0x6C,
            0x00,0x69,0x00,0x62,0x00,0x72,0x00,0x69,
            0x00,

            //bolded font?
             0x2B,0x27,0xDC,0x00,0x01,0x00,0xBC,0x02,
            0x00,0x00,0x00,0x02,0xEE,0x00,0x07,0x01,
            0x00,0x00,0x00,0x00,0x00,0xFF,0x02,0x07,
            0x00,0x00,0x00,0x43,0x00,0x61,0x00,0x6C,
            0x00,0x69,0x00,0x62,0x00,0x72,0x00,0x69,
            0x00,

            0x25,0x06,0x01,0x00,0x02,0x0E,0x00,0x80,0x81,0x08,0x00,0x26,0x00,0xE4,0x04,0x00,0xDB,0x04,0x04,0x02,0x00,0x00,0x00,
            0x2D,0x44,0x00,0x00,0x00,0x00,0x03,0x40,0x00,0x00,0x00,0x00,0x00,0xFF,0x03,0x41,0x00,0x00,0xFF,0xFF,0xFF,0xFF,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x2D,0x44,0x11,0x00,0x00,0x00,0x03,0x40,0x00,0x00,0x00,0x00,0x00,0xFF,0x03,0x41,0x00,0x00,0xFF,0xFF,0xFF,0xFF,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0xDC,0x04,0x00,0xE5,0x04,0x04,0x01,0x00,0x00,0x00,0x2E,0x33,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xE6,0x04,0x00,0xF2,
            0x04,0x04,0x01,0x00,0x00,0x00,0x2F,0x10,0xFF,0xFF,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,
            0x00,0xF3,0x04,0x00,0xE9,0x04,0x04,
            
            0x04,
            0x00,0x00,0x00,
                                         //(#font)
            0x2F,0x10,0x00,0x00,0x00,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,0x00,// standard 
            0x2F,0x10,0x00,0x00,0xA4,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x01,0x00,// datetime
            0x2F,0x10,0x00,0x00,0xA6,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x01,0x00,//date
            0x2F,0x10,0x00,0x00,0x00,0x01,    0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,0x00,//standard bolded

            0xEA,0x04,0x00,0xEB,0x04,0x04,0x01,0x00,

            0x00,0x00,0x25,0x06,0x01,0x00,0x02,0x11,0x00,0x80,0x80,0x18,0x10,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x26,0x00,0x30,0x1C,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x08,0x00,0x00,0x00,0x4E,
            0x00,0x6F,0x00,0x72,0x00,0x6D,0x00,0x61,0x00,0x6C,0x00,0x6E,0x00,0x79,0x00,0xEC,0x04,0x00,0xF9,0x03,0x04,0x00,0x00,
            0x00,0x00,0xFA,0x03,0x00,0xFC,0x03,0x50,0x00,0x00,0x00,0x00,0x11,0x00,0x00,0x00,0x54,0x00,0x61,0x00,0x62,0x00,0x6C,
            0x00,0x65,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4D,0x00,0x65,0x00,0x64,0x00,0x69,0x00,0x75,0x00,
            0x6D,0x00,0x32,0x00,0x11,0x00,0x00,0x00,0x50,0x00,0x69,0x00,0x76,0x00,0x6F,0x00,0x74,0x00,0x53,0x00,0x74,0x00,0x79,
            0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,0x67,0x00,0x68,0x00,0x74,0x00,0x31,0x00,0x36,0x00,0xFD,0x03,0x00,0x23,
            0x04,0x02,0x0E,0x00,0x00,0xEB,0x08,0x00,0xF6,0x08,0x2A,0x00,0x00,0x00,0x00,0x11,0x00,0x00,0x00,0x53,0x00,0x6C,0x00,
            0x69,0x00,0x63,0x00,0x65,0x00,0x72,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,0x67,
            0x00,0x68,0x00,0x74,0x00,0x31,0x00,0xF7,0x08,0x00,0xEC,0x08,0x00,0x24,0x00,0x23,0x04,0x03,0x0F,0x00,0x00,0xB0,0x10,
            0x00,0xB2,0x10,0x32,0x00,0x00,0x00,0x00,0x15,0x00,0x00,0x00,0x54,0x00,0x69,0x00,0x6D,0x00,0x65,0x00,0x53,0x00,0x6C,
            0x00,0x69,0x00,0x63,0x00,0x65,0x00,0x72,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,
            0x67,0x00,0x68,0x00,0x74,0x00,0x31,0x00,0xB3,0x10,0x00,0xB1,0x10,0x00,0x24,0x00,0x97,0x02,0x00
        ];

        private readonly static byte[] _workbookBinStart =
        [
            0x83,0x01,0x00,0x80,0x01,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x02,0x00,0x00,0x00,0x78,0x00,0x6C,0x00,0x01,0x00,0x00,0x00,
            0x37,0x00,0x01,0x00,0x00,0x00,0x36,0x00,0x05,0x00,0x00,0x00,0x32,0x00,0x34,0x00,0x33,0x00,0x32,0x00,0x36,0x00,0x99,0x01,0x0C,0x20,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x25,0x06,0x01,0x00,0x03,0x0F,0x00,0x80,0x97,0x10,0x34,0x18,0x00,0x00,0x00,0x43,0x00,0x3A,0x00,0x5C,0x00,0x73,0x00,0x71,0x00,0x6C,0x00,0x73,0x00,0x5C,0x00,
            0x54,0x00,0x65,0x00,0x73,0x00,0x74,0x00,0x79,0x00,0x5A,0x00,0x61,0x00,0x70,0x00,0x69,0x00,0x73,0x00,0x75,0x00,0x58,0x00,0x6C,0x00,0x73,0x00,0x62,0x00,0x5C,0x00,0x26,0x00,
            0x25,0x06,0x01,0x00,0x00,0x10,0x00,0x80,0x81,0x18,0x82,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x2F,0x00,0x00,0x00,0x31,0x00,0x33,0x00,0x5F,0x00,0x6E,0x00,0x63,0x00,
            0x72,0x00,0x3A,0x00,0x31,0x00,0x5F,0x00,0x7B,0x00,0x31,0x00,0x36,0x00,0x35,0x00,0x30,0x00,0x38,0x00,0x44,0x00,0x36,0x00,0x39,0x00,0x2D,0x00,0x43,0x00,0x46,0x00,0x38,0x00,
            0x37,0x00,0x2D,0x00,0x34,0x00,0x37,0x00,0x36,0x00,0x39,0x00,0x2D,0x00,0x38,0x00,0x34,0x00,0x35,0x00,0x36,0x00,0x2D,0x00,0x44,0x00,0x34,0x00,0x41,0x00,0x34,0x00,0x30,0x00,
            0x31,0x00,0x31,0x00,0x33,0x00,0x31,0x00,0x35,0x00,0x36,0x00,0x37,0x00,0x7D,0x00,0x2F,0x00,0x00,0x00,0x2F,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x00,0x00,0x00,0x26,0x00,0x87,0x01,0x00,0x25,0x06,0x01,0x00,0x02,0x10,0x00,0x80,0x80,0x18,0x10,0x00,0x00,0x00,0x00,0x0D,0x00,0x00,0x00,0xFF,0xFF,0xFF,0xFF,
            0x00,0x00,0x00,0x00,0x26,0x00,0x9E,0x01,0x1D,0x00,0x00,0x00,0x00,0x9E,0x16,0x00,0x00,0xB4,0x69,0x00,0x00,0xE8,0x26,0x00,0x00,0x58,0x02,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
            0x00,0x00,0x00,0x78,0x88,0x01,0x00,
            0x8F,0x01,0x00
        ];

        private readonly static byte[] _workbookBinMiddle =
        [
            0x90,0x01,0x00
        ];

        private readonly static byte[] _workbookBinEnd =
        [
            0x9D,0x01,0x1A,0x35,0xEA,0x02,0x00,0x01,0x00,0x00,0x00,0x64,0x00,0x00,0x00,0xFC,0xA9,0xF1,0xD2,0x4D,0x62,0x50,0x3F,0x01,
            0x00,0x00,0x00,0x6A,0x00,0x9B,0x01,0x01,0x00,0x23,0x04,0x03,0x0F,0x00,0x00,0xAB,0x10,0x01,0x01,0x24,0x00,0x84,0x01,0x00
        ];

        private readonly static byte[] _binaryIndexBin =
        [
            0x2A,0x18,0x00,0x00,0x00,0x00,0x20,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x95,
            0x02,0x00
        ];

        private BufferedStream _stream;

        private int _columnCount;
        private uint _startCol;
        private uint _endCol;
        private byte[] _colA;
        private byte[] _colZ;

        private const int _rRkIntegerLowerLimit = -1 << 29;
        private const int _rRkIntegerUpperLimit = (1 << 29) - 1;

        private readonly CompressionLevel _compressionLevel = CompressionLevel.Fastest;
        public XlsbWriter(string path, CompressionLevel compressionLevel = CompressionLevel.Fastest) 
            : this(new FileStream(path, FileMode.Create), compressionLevel, leaveExcelArchiveOpen: false)
        {
            _excelStreamWasProvided = false;
        }

        public XlsbWriter(Stream stream, CompressionLevel compressionLevel = CompressionLevel.Fastest, bool leaveExcelArchiveOpen = true)
        {
            _excelStreamWasProvided = true;
            sheetCnt = 0;
            _compressionLevel = compressionLevel;
            try
            {
                _newExcelFileStream = stream;
                _excelArchiveFile = new ZipArchive(_newExcelFileStream, ZipArchiveMode.Create, leaveExcelArchiveOpen);
            }
            catch (Exception)
            {
                throw new Exception("creation file error");
            }
        }

        public override void AddSheet(string sheetName, bool hidden = false)
        {
            sheetCnt++;
            _sheetList.Add((sheetName, $"xl/worksheets/sheet{sheetCnt}.bin", null, hidden, $"sheet{sheetCnt}.bin", sheetCnt,null));
        }

        internal record FilterData
        {
            public byte SheetIndex { get; set; }
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
            public int StartRow { get; set; }
            public int EndRow { get; set; }
        }

        private List<FilterData> _filteredDict;

        public override void WriteSheet(IDataReader dataReader, Boolean headers = true, int overLimit = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            if (doAutofilter)
            {
                _autofilterIsOn = true;
                _filteredDict ??= new();
            }
            this._areHeaders = headers;
            _dataColReader = new DataColReader(dataReader, headers, overLimit);

            int rowNum = 0;
            _columnCount = _dataColReader.FieldCount;

            _startCol = (uint)startingColumn;
            _endCol = (uint)(_startCol + _columnCount);

            _colWidesArray = new double[_columnCount];
            Array.Fill<double>(_colWidesArray, -1.0);

            typesArray = new int[_columnCount];
            _newTypes = new TypeCode[_columnCount];

            var rdr = _dataColReader._dataReader;
            for (int l = 1; l <= _columnCount; l++)
            {
                int lenn = rdr.GetName(l - 1).Length + (doAutofilter?2:0);
                double tempWidth = 1.25 * lenn + 2;
                if (tempWidth > _MAX_WIDTH)
                {
                    tempWidth = _MAX_WIDTH;
                }
                if (_colWidesArray[l - 1] < tempWidth)
                {
                    _colWidesArray[l - 1] = tempWidth;
                }
            }

            int nr = 0;
            _dataColReader.top100 = new List<object[]>();
            bool areNextRows = true;
            while (nr < 100)
            {
                areNextRows = rdr.Read();
                if (!areNextRows)
                {
                    break;
                }

                object[] arr = new object[rdr.FieldCount];
                rdr.GetValues(arr);

                for (int i = 0;i< arr.Length;i++)
                {
                    if (arr[i] is Memory<byte> mem)
                    {
                        Memory<byte> m = new byte[mem.Length];
                        mem.CopyTo(m);
                        arr[i] = m;
                    }
                }

                _dataColReader.top100.Add(arr);
                nr++;
                SetColsLengtth(_columnCount, arr);
            }
            areNextRows = rdr.Read();
            _dataColReader.AreNextRows = areNextRows;

            if (sheetCnt != 1)
            {
                _sheet1Bytes[54] = 0x9C; // only first is selected
            }
            var newEntry = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt - 1].pathInArchive, _compressionLevel);
            _stream = new BufferedStream(newEntry.Open());
            try
            {
                InitSheet(doAutofilter); //lock first row
                while (_dataColReader.Read())
                {
                    if (rowNum == 0 || _areHeaders && rowNum == 1)
                    {
                        if (rowNum == 0 && _areHeaders)
                        {
                            for (int i = 0; i < _columnCount; i++)
                            {
                                typesArray[i] = 0;
                                _newTypes[i] = TypeCode.String;
                            }
                        }
                        else
                        {
                            ExcelWriter.SetTypes(_dataColReader, typesArray, _newTypes, _columnCount, detectBoolenaType: true);
                        }
                    }

                    InitRow(rowNum);
                    WriteRow(rowNum==0);

                    if (rowNum % 10000 == 0)
                    {
                        DoOn10k(rowNum);
                    }
                    rowNum++;
                }
                _rowsCount = rowNum - 1;
                
                _stream.Write(_sheet1Bytes[218..290].AsSpan());

                if (doAutofilter)
                {
                    Span<byte> buff = stackalloc byte[8];
                    _stream.Write(_autoFilterStartBytes);
                    Int32ToSpecificBuffer(buff, startingRow, 0);
                    Int32ToSpecificBuffer(buff, startingRow + _rowsCount, 4);
                    _stream.Write(buff);
                    Int32ToSpecificBuffer(buff, startingColumn, 0);
                    Int32ToSpecificBuffer(buff, startingColumn + _columnCount - 1, 4);
                    _stream.Write(buff);
                    _stream.Write(_autoFilterEndBytes);

                   _filteredDict.Add(new FilterData()
                    {
                       SheetIndex = (byte)(_sheetList.Count - 1),
                        StartColumn = startingColumn,
                        EndColumn = startingColumn + _columnCount - 1,
                        StartRow = startingRow,
                        EndRow = startingRow + _rowsCount,
                    });
                }

                _stream.Write(_sheet1Bytes[290..].AsSpan());

                //stream.Write(new byte[] 
                //{ 
                //    /*0x25 ,0x06, 0x01, 0x00, 0x00, 0x10, 0x00, 0x80
                //    , 0x80 , 0x18 , 0x10 , 0x00 , 0x00 , 0x00 , 0x00 , 0x01
                //    , 0x00 , 0x00 , 0x00 , 0x00 , 0x00 , 0x00 , 0x00 , 0x00
                //    , 0x00 , 0x00 , 0x00 , 0x26 , 0x00 ,*/ 0xA1 , 0x01 , 0x10
                //    , 0x00 , 0x00 , 0x00 , 0x00 , 0x14 , 0x00 , 0x00 , 0x00
                //    , 0x00 , 0x00 , 0x00 , 0x00 , 0x03 , 0x00 , 0x00 , 0x00
                //    , 0xA2 , 0x01 , 0x00
                //});


                //stream.Write(sheet1Bytes, 218, sheet1Bytes.Length - 218); // całkowity koniec
            }
            finally
            {
                _stream.Dispose();
            }
            //throw new NotImplementedException();
        }

        private void WriteRow(bool boldedStyle)
        {
            for (int column = 0; column < _columnCount; column++)
            {
                if (_dataColReader.IsDBNull(column))
                    continue;

                if (_newTypes[column] == TypeCode.String) // string
                {
                    string stringValue = _dataColReader.GetString(column);
                    WriteString(stringValue, column, boldedStyle);
                }
                else if (typesArray[column] == 5) // Memory<byte>
                {
                    var stringValue = Encoding.UTF8.GetString(((Memory<byte>)(_dataColReader.GetValue(column))).Span);
                    WriteString(stringValue, column, boldedStyle);
                }
                else if(_newTypes[column] == TypeCode.Object)
                {
                    string stringValue = _dataColReader.GetValue(column).ToString();
                    WriteString(stringValue, column);
                }
                else if (_newTypes[column] == TypeCode.Boolean) // bool
                {
                    WriteBool(_dataColReader.GetBoolean(column), column);
                }
                else if (typesArray[column] == 1)//number
                {

                    switch (_newTypes[column])
                    {
                        case TypeCode.Byte:
                            byte byteValue = _dataColReader.GetByte(column);
                            WriteRkNumberInteger(byteValue, column);
                            break;
                        case TypeCode.SByte:
                            sbyte sbyteValue = _dataColReader.GetSByte(column);
                            WriteRkNumberInteger(sbyteValue, column);
                            break;
                        case TypeCode.Int16:
                            Int16 int16Value = _dataColReader.GetInt16(column);
                            WriteRkNumberInteger(Convert.ToInt32(int16Value), column);
                            break;
                        case TypeCode.Int32:

                            Int32 int32Value = _dataColReader.GetInt32(column);
                            if (int32Value >= _rRkIntegerLowerLimit && int32Value <= _rRkIntegerUpperLimit)
                            {
                                WriteRkNumberInteger(int32Value, column);
                            }
                            else
                            {
                                WriteDouble((double)int32Value, column);
                            }
                            break;
                        case TypeCode.Int64:
                            Int64 int64Value = _dataColReader.GetInt64(column);
                            WriteDouble(Convert.ToDouble(int64Value), column);
                            break;
                        case TypeCode.Single:
                            float floatVal = _dataColReader.GetFloat(column);
                            WriteDouble(Convert.ToDouble(floatVal), column);
                            break;
                        case TypeCode.Double:
                            double doubleVal = _dataColReader.GetDouble(column);
                            WriteDouble(doubleVal, column);
                            break;
                        case TypeCode.Decimal:
                            decimal decimalVal = _dataColReader.GetDecimal(column);
                            WriteDouble(decimal.ToDouble(decimalVal), column);
                            break;
                        default:
                            throw new Exception("number format Excel error");
                    }
                }
                else if (typesArray[column] == 2) //date
                {
                    DateTime dtVal = _dataColReader.GetDateTime(column);
                    WriteDate(dtVal, column);
                }
                else if (typesArray[column] == 3) //dateTime
                {
                    DateTime dtVal = _dataColReader.GetDateTime(column);
                    if (SuppressSomeDate && (dtVal as DateTime?).Value.Year == 1000)//1000-xx-xx
                    {
                        continue;
                    }
                    WriteDateTime(dtVal, column);
                }
                
            }
        }

        public override void WriteSheet(string[] oneColumn)
        {
            if (sheetCnt != 1)
            {
                _sheet1Bytes[54] = 156; // only first is selected
            }
            var newEntry = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt - 1].pathInArchive, _compressionLevel);
            _stream = new BufferedStream(newEntry.Open());
            try
            {
                InitSheet(false);
                for (int rowNum = 0; rowNum < oneColumn.Length; rowNum++)
                {
                    string txt = oneColumn[rowNum];
                    InitRow((int)rowNum);
                    WriteString(txt, 0);
                }

                _stream.Write(_sheet1Bytes, 218, _sheet1Bytes.Length - 218); // całkowity koniec
            }
            finally
            {
                _stream.Dispose();
            }
        }

        public override void Dispose()
        {
            DoOnCompress();
            Save();
        }

        private void WriteColsWidth()
        {
            //widths !!!
            _stream.WriteByte(134);
            _stream.WriteByte(3);
            int l = 0;
            for (uint i = _startCol; i < _endCol; i++)
            {
                // start of column definition   
                _stream.WriteByte(0);
                _stream.WriteByte(60);
                _stream.WriteByte(18);
                //column min
                _stream.Write(BitConverter.GetBytes(i));
                // column max
                _stream.Write(BitConverter.GetBytes(i));
                //width
                _stream.WriteByte(0);
                _stream.WriteByte((byte)(_colWidesArray[l])); // .. x 7 = pixels
                _stream.WriteByte(0);
                _stream.WriteByte(0);

                _stream.WriteByte(0);
                _stream.WriteByte(0);
                _stream.WriteByte(0);
                _stream.WriteByte(0);
                _stream.WriteByte(2); // column properties /hidden etc, 2 = normal
                // end of column definition   
                l++;
            }
            _stream.WriteByte(0);
            _stream.WriteByte(135);
            _stream.WriteByte(3);
            _stream.WriteByte(0);
        }

        private void InitSheet(bool doAutofiler)
        {
            _colA = BitConverter.GetBytes(_startCol); // start col
            _colZ = BitConverter.GetBytes(_endCol); // end col

            _colA.CopyTo(_sheet1Bytes, 40);
            _colZ.CopyTo(_sheet1Bytes, 44);

            _stream.Write(_sheet1Bytes.AsSpan()[0..84]); // start of file
            if (doAutofiler)
            {
                _stream.Write(_stickHeaderA1bytes);
            }

            _stream.Write(_sheet1Bytes.AsSpan()[84..159]); // start of file

            WriteColsWidth();
            _stream.Write(_sheet1Bytes, 159, 175 - 159); // BrtACBegin
            _stream.WriteByte(38); // pos. 175 ?
            _stream.WriteByte(0); // pos. 176 // BrtACEnd
        }

        //private readonly static byte[] rowNeededBytes = { 0, 0, 0, 0, 44, 1, 0, 0, 0, 1, 0, 0, 0 };
        private void InitRow(int rowNumber)
        {
            Span<byte> buff = stackalloc byte[27];
            //buff[0] = 0; stackalloct is 0,0..
            buff[1] = 25;
            BitConverter.TryWriteBytes(buff[2..], (int)rowNumber);
            buff[10] = 44;
            buff[11] = 1;
            buff[15] = 1;
            BitConverter.TryWriteBytes(buff[(6 + 13)..], (int)_startCol);
            BitConverter.TryWriteBytes(buff[(6 + 13 + 4)..], (int)_endCol);
            
            _stream.Write(buff);// 6 + 13 + 4 + 4 = 27
        }

        private void WriteDouble(double val, int colNum/*, int offset = 0*/, byte styleNum = 0)
        {
            Span<byte> buff = stackalloc byte[18];
            buff[0] = 5;
            buff[1] = 16;//8+8
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);
            buff[6] = styleNum;
            //_buffer[7] = 0;
            //_buffer[8] = 0;
            //_buffer[9] = 0;
            BitConverter.TryWriteBytes(buff[10..], val);
            _stream.Write(buff);
        }

        private void WriteBool(bool val, int colNum)
        {
            Span<byte> buff = stackalloc byte[11];
            buff[0] = 0x04;
            buff[1] = 8 + 1;
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);
            //style
            //generalStyle.CopyTo(_buffer, 6); generalStyle = [0,0,0,0]
            buff[10] = (byte)(val ? 1 : 0); // 0 = false, 1 = true
            //buff[11] = 1;
            _stream.Write(buff);
        }

        private void WriteRkNumberInteger(int val, int colNum/*, int offset = 0*/, byte styleNum = 0)
        {
            Span<byte> buff = stackalloc byte[14];
            buff[0] = 2;
            buff[1] = 12;//8+4
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);
            buff[6] = styleNum;
            //buff[7] = 0;
            //buff[8] = 0;
            //buff[9] = 0; stackalloc set bytes to 0

            val <<= 2;
            val |= 0b00000010; // = integer flag

            BitConverter.TryWriteBytes(buff[10..], val);
            _stream.Write(buff);
        }

        // 
        //private void writeIntegerRkGeneralNumber(int val, int colNum/*, int offset = 0*/, byte styleNum = 0)
        //{
        //    //stream.WriteByte(2); // 2= rknumber 5 = double, pos 205, 7 shared string
        //    //stream.WriteByte(8 + 4); // int  , 4 column, 4 style, 4/8 = number
        //    //stream.Write(BitConverter.GetBytes(colNum)); // column number
        //    //stream.Write(BitConverter.GetBytes((uint)0)); // style number stylu 0 - general, 1 date time, 2 date
        //    //Int32RkNumWrite(val, 0);
        //    //stream.Write(_buffer,0, 4);

        //    _buffer[/*offset*/ +0] = 2;
        //    _buffer[/*offset*/ +1] = 12;//8+4
        //    Int32ToBuffer(colNum, /*offset*/ +2);
        //    _buffer[/*offset*/ +6] = styleNum;
        //    _buffer[/*offset*/ +7] = 0;
        //    _buffer[/*offset*/ +8] = 0;
        //    _buffer[/*offset*/ +9] = 0;
        //    RkNumberGeneralWrite((double)val, /*offset*/ +10, false);
        //    stream.Write(_buffer, /*offset*/ +0, /*offset*/ +14);
        //}


        private void WriteString(string stringValue, int colNum, bool bolded = false)
        {
            ref var index = ref CollectionsMarshal.GetValueRefOrAddDefault(_sstDic, stringValue, out bool exists);
            if (!exists)
            {
                index = _sstCntUnique;
                _sstCntUnique++;
            }
            WriteStringFromShared(index, colNum, bolded);
            _sstCntAll++;
        }

        private void WriteDateTime(DateTime dateTime, int colNum)
        {
            double d1 = dateTime.ToOADate();
            WriteDouble(d1, colNum, 1); // 1 = datetime
        }
        private void WriteDate(DateTime dateTime, int colNum)
        {
            double d1 = dateTime.ToOADate();
            if (d1 > 1 << 20 - 50_000 || d1 < -(1 << 20 - 50_000))
            {
                WriteDouble(d1, colNum, 2); // 2 = date
            }
            else
            {
                Span<byte> buff = stackalloc byte[14];
                buff[0] = 2;
                buff[1] = 12;//8+4
                BitConverter.TryWriteBytes(buff[2..], (int)colNum);
                //generalStyle.CopyTo(_buffer, /*offset*/ + 6);
                buff[6] = 2;
                buff[7] = 0;
                buff[8] = 0;
                buff[9] = 0;
                RkNumberGeneralWrite(buff[10..],d1);
                _stream.Write(buff);
            }
        }

        private void WriteStringFromShared(int val, int colNum, bool bolded = false)
        {
            Span<byte> buff = stackalloc byte[14];
            buff[0] = 7;
            buff[1] = 12;//8+4
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);
            //generalStyle.CopyTo(buff, 6); generalStyle = [0,0,0,0]

            buff[6] = (byte)(bolded ? 3 : 0);


            BitConverter.TryWriteBytes(buff[10..], (int)val);
            _stream.Write(buff);
        }

        internal override void FinalizeFile()
        {
            SaveSst();
            var newEntry = _excelArchiveFile.CreateEntry(@"xl/styles.bin", _compressionLevel);
            using (var str = newEntry.Open())
            {
                using var sw = new BinaryWriter(str);
                sw.Write(_stylesBin);
            }

            newEntry = _excelArchiveFile.CreateEntry(@"xl/workbook.bin", _compressionLevel);
            using (var str = newEntry.Open())
            {
                using var sw = new BinaryWriter(str);
                sw.Write(_workbookBinStart);

                for (int i = 0; i < _sheetList.Count; i++)
                {
                    var (name, pathInArchive, pathOnDisc, isHidden, nameInArchive, sheetId, _) = _sheetList[i];

                    string rId = $"rId{sheetId}";

                    sw.Write((byte)156);
                    sw.Write((byte)1);
                    sw.Write((byte)(4 + 3 * 4 + name.Length * 2 + rId.Length * 2));

                    if (isHidden)
                        sw.Write(BitConverter.GetBytes((int)1));
                    else
                        sw.Write(BitConverter.GetBytes((int)0));


                    sw.Write(BitConverter.GetBytes(sheetId));
                    sw.Write(BitConverter.GetBytes(rId.Length));
                    foreach (var m in rId)
                    {
                        sw.Write((byte)(m));
                        sw.Write((byte)(m >> 8));
                    }
                    sw.Write(BitConverter.GetBytes(name.Length));
                    foreach (var m in name)
                    {
                        sw.Write((byte)(m));
                        sw.Write((byte)(m >> 8));
                    }
                }


                sw.Write(_workbookBinMiddle);
                WriteFilterDefinedNames(sw);

                sw.Write(_workbookBinEnd);
            }

            for (int i = 0; i < _sheetList.Count; i++)
            {
                var (_, _, _, _, _, sheetId,_) = _sheetList[i];
                newEntry = _excelArchiveFile.CreateEntry($@"xl/worksheets/binaryIndex{sheetId}.bin", _compressionLevel);
                using var str = newEntry.Open();
                using var sw = new BinaryWriter(str);
                sw.Write(_binaryIndexBin);
            }

            newEntry = _excelArchiveFile.CreateEntry(@"[Content_Types].xml", _compressionLevel);
            using (var str = newEntry.Open())
            {
                using var sw = new StreamWriter(str, Encoding.UTF8);
                sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                sw.Write(@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">");
                sw.Write(@"<Default Extension=""bin"" ContentType=""application/vnd.ms-excel.sheet.binary.macroEnabled.main""/>");
                sw.Write(@"<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>");
                sw.Write(@"<Default Extension=""xml"" ContentType=""application/xml""/>");

                for (int i = 0; i < _sheetList.Count; i++)
                {
                    var (name, pathInArchive, pathOnDisc, isHidden, nameInArchive, sheetId,_) = _sheetList[i];

                    sw.Write($@"<Override PartName=""/{pathInArchive}"" ContentType=""application/vnd.ms-excel.worksheet""/>");
                    sw.Write($@"<Override PartName=""/xl/worksheets/binaryIndex{sheetId}.bin"" ContentType=""application/vnd.ms-excel.binIndexWs""/>");
                }

                sw.Write(@"<Override PartName=""/xl/styles.bin"" ContentType=""application/vnd.ms-excel.styles""/>");
                sw.Write(@"<Override PartName=""/xl/sharedStrings.bin"" ContentType=""application/vnd.ms-excel.sharedStrings""/>");

                if (!String.IsNullOrWhiteSpace(DocPopertyProgramName))
                {
                    sw.Write(@"<Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/>");
                    sw.Write(@"<Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml""/>");
                }
                sw.Write(@"</Types>");
            }

            newEntry = _excelArchiveFile.CreateEntry($"xl/_rels/workbook.bin.rels", _compressionLevel);
            using (var str = newEntry.Open())
            {
                using var sw = new StreamWriter(str, Encoding.UTF8);
                sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                sw.Write(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">");

                for (int i = 0; i < _sheetList.Count; i++)
                {
                    var (name, pathInArchive, pathOnDisc, isHidden, nameInArchive, sheetId,_) = _sheetList[i];
                    string rId = $"rId{sheetId}";
                    sw.Write($@"<Relationship Id=""{rId}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/{nameInArchive}""/>");
                }

                sw.Write($@"<Relationship Id=""rId{_sheetList.Count + 2}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.bin""/>");
                sw.Write($@"<Relationship Id=""rId{_sheetList.Count + 3}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"" Target=""sharedStrings.bin""/>");
                sw.Write(@"</Relationships>");
            }

            newEntry = _excelArchiveFile.CreateEntry($"_rels/.rels", _compressionLevel);
            using (var str = newEntry.Open())
            {
                using var sw = new StreamWriter(str, Encoding.UTF8);
                sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                sw.Write(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">");
                if (!String.IsNullOrWhiteSpace(DocPopertyProgramName))
                {
                    sw.Write(@"<Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml""/>");
                    sw.Write(@"<Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/>");
                }
                sw.Write(@"<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.bin""/>");
                sw.Write(@"</Relationships>");
            }

            if (!String.IsNullOrWhiteSpace(DocPopertyProgramName))
            {
                newEntry = _excelArchiveFile.CreateEntry($"docProps/app.xml", _compressionLevel);
                using (var str = newEntry.Open())
                {
                    using var sw = new StreamWriter(str, Encoding.UTF8);
                    sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                    sw.Write(@"<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">");
                    sw.Write($"<Application>{DocPopertyProgramName}</Application>");
                    sw.Write(@"<DocSecurity>0</DocSecurity>");
                    sw.Write(@"<ScaleCrop>false</ScaleCrop>");
                    sw.Write(@"<HeadingPairs><vt:vector size=""2"" baseType=""variant""><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>");
                    sw.Write($"<vt:variant><vt:i4>{_sheetList.Count}</vt:i4></vt:variant></vt:vector></HeadingPairs>");
                    sw.Write("<TitlesOfParts>");
                    sw.Write($"<vt:vector size=\"{_sheetList.Count}\" baseType=\"lpstr\">");
                    foreach (var (name, _, _, _, _, _, _) in _sheetList)
                    {
                        sw.Write($"<vt:lpstr>{name}</vt:lpstr>");
                    }
                    sw.Write($@"</vt:vector>");
                    sw.Write($@"</TitlesOfParts>");
                    sw.Write(@"<Company></Company>");
                    sw.Write(@"<LinksUpToDate>false</LinksUpToDate>");
                    sw.Write(@"<SharedDoc>false</SharedDoc>");
                    sw.Write(@"<HyperlinksChanged>false</HyperlinksChanged>");
                    sw.Write(@"<AppVersion>16.0300</AppVersion>");
                    sw.Write(@"</Properties>");
                }

                newEntry = _excelArchiveFile.CreateEntry($"docProps/core.xml", _compressionLevel);
                using (var str = newEntry.Open())
                {
                    using var sw = new StreamWriter(str, Encoding.UTF8);
                    sw.WriteLine($@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                    sw.Write($@"<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">");
                    sw.Write($@"<dc:creator>{DocPopertyProgramName} - used by {Environment.UserName}</dc:creator>");
                    sw.Write($@"<cp:lastModifiedBy>{DocPopertyProgramName} - used by {Environment.UserName}</cp:lastModifiedBy>");
                    sw.Write($@"<dcterms:created xsi:type=""dcterms:W3CDTF"">2015-06-05T18:19:34Z</dcterms:created>");
                    sw.Write($@"<dcterms:modified xsi:type=""dcterms:W3CDTF"">2021-09-05T11:11:46Z</dcterms:modified>");
                    sw.Write($@"</cp:coreProperties>");
                }
            }

            foreach (var (_, _, _, _, nameInArchive, sheetId, _) in _sheetList)
            {
                newEntry = _excelArchiveFile.CreateEntry($"xl/worksheets/_rels/{nameInArchive}.rels", _compressionLevel);
                using var str = newEntry.Open();
                using var sw = new StreamWriter(str, Encoding.UTF8);
                sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                sw.Write(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">");
                sw.Write($@"<Relationship Id=""rId1"" Type=""http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex"" Target=""binaryIndex{sheetId}.bin""/>");
                sw.Write(@"</Relationships>");
            }

        }

        private static readonly byte[] _magicFilterExcel2016Fix0 = [0xE1, 0x02, 0x00, 0xE5, 0x02, 0x00, 0xEA, 0x02];
        private static readonly byte[] _magicFilterExcel2016Fix1 = [
            0x27,
            0x46,
            0x21,
            0x00,
            0x00,
            0x00,
            0x00,
            255,// -> (byte)sheetIndex,
            0x00,
            0x00,
            0x00,
            0x0F,
            0x00,
            0x00,
            0x00,
            0x5F,
            0x00,   // _0, // FilterDatabase (UTF16) - starts
            0x46,
            0x00,
            0x69,
            0x00,
            0x6C,
            0x00,
            0x74,
            0x00,
            0x65,
            0x00,
            0x72,
            0x00,
            0x44,
            0x00,
            0x61,
            0x00,
            0x74,
            0x00,
            0x61,
            0x00,
            0x62,
            0x00,
            0x61,
            0x00,
            0x73,
            0x00,
            0x65,
            0x00,// FilterDatabase (UTF16) - ends
            0x0F,
            0x00,
            0x00,
            0x00,
            0x3B,
            255,//->(byte)sheetNum,
            0x00
        ];
        private static readonly byte[] _magicFilterExcel2016Fix2 = [0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF];
        
        private void WriteFilterDefinedNames(BinaryWriter sw)
        {
            int? filteredDictIntemsCnt = _filteredDict?.Count;
            //temportary fix for https://github.com/KrzysztofDusko/SpreadSheetTasks/issues/2
            if (_autofilterIsOn && filteredDictIntemsCnt >= 0 && (0x80 + (filteredDictIntemsCnt - 21) * 0x0c) <= Byte.MaxValue)
            {
                sw.Write(_magicFilterExcel2016Fix0);
                if (filteredDictIntemsCnt <= 10)
                {
                    sw.Write([(byte)(0x10 + (filteredDictIntemsCnt - 1) * 0x0c), (byte)filteredDictIntemsCnt]);// !!! ? for cnt <=10
                    sw.Write([0x00, 0x00, 0x00]);
                }
                else if (filteredDictIntemsCnt <= 20)
                {
                    sw.Write([(byte)(0x10 + (filteredDictIntemsCnt - 1) * 0x0c), (byte)((filteredDictIntemsCnt - 1) / 10), (byte)filteredDictIntemsCnt]);
                    sw.Write([0x00, 0x00, 0x00]);
                }
                else // ???
                {
                    sw.Write([(byte)(0x80 + (filteredDictIntemsCnt - 21) * 0x0c), (byte)((filteredDictIntemsCnt - 1) / 10), (byte)filteredDictIntemsCnt]);
                    sw.Write([0x00, 0x00, 0x00]);
                }

                for (int nm = 0; nm < filteredDictIntemsCnt; nm++)
                {
                    sw.Write([0x00, 0x00, 0x00, 0x00]);
                    sw.Write([(byte)(_filteredDict[nm].SheetIndex), 0x00, 0x00, 0x00]);
                    sw.Write([(byte)(_filteredDict[nm].SheetIndex), 0x00, 0x00, 0x00]);
                }

                sw.Write([0xE2, 0x02, 0x00]);

                for (int sheetNum = 0; sheetNum < _filteredDict.Count; sheetNum++)
                {
                    int startColumn = _filteredDict[sheetNum].StartColumn;
                    int endColumn = _filteredDict[sheetNum].EndColumn;
                    int startRow = _filteredDict[sheetNum].StartRow;
                    int endRow = _filteredDict[sheetNum].EndRow;
                    byte sheetIndex = _filteredDict[sheetNum].SheetIndex;
                    _magicFilterExcel2016Fix1[7] = (byte)sheetIndex;
                    _magicFilterExcel2016Fix1[^2] = (byte)sheetNum;


                    sw.Write(_magicFilterExcel2016Fix1);
                    sw.Write(BitConverter.GetBytes(startRow));
                    sw.Write(BitConverter.GetBytes(endRow));
                    sw.Write(BitConverter.GetBytes((Int16)startColumn));
                    sw.Write(BitConverter.GetBytes((Int16)endColumn));
                    sw.Write(_magicFilterExcel2016Fix2);
                }
            }
        }

        private readonly static byte[] _startSst = [159, 1, 8];// SharedStringStart = 159
        private readonly static byte[] _endSst = [160, 1, 0]; // SharedStringEnd = 160

        private void SaveSst()
        {
            var newSST = _excelArchiveFile.CreateEntry($"xl/sharedStrings.bin", _compressionLevel);
            int bufferLen = 1 << 17;// 2*256 * 256;
            var _buffX = ArrayPool<byte>.Shared.Rent(bufferLen);
            try
            {
                using var sstStream = newSST.Open();
                using var binaryWriter = new BinaryWriter(sstStream);
                //SST
                binaryWriter.Write(_startSst);
                binaryWriter.Write((int)_sstCntUnique);
                binaryWriter.Write((int)_sstCntAll);

                int localPosition = 0;
                foreach (var txt in _sstDic.Keys)
                {
                    int txtLen = txt.Length;
                    if (txtLen > 32767)
                    {
                        txtLen = 32767;
                    }

                    int recLen = 5 + 2 * txtLen;
                    //localPosition = 0;
                    if (localPosition + recLen + 10 > bufferLen)
                    {
                        binaryWriter.Write(_buffX, 0, localPosition);
                        localPosition = 0;
                    }

                    _buffX[localPosition++] = 19;

                    if (recLen >= 128)
                    {
                        _buffX[localPosition++] = (byte)(128 + (recLen % 128));
                        int tmp = recLen >> 7;
                        if (tmp >= 256)
                        {
                            _buffX[localPosition++] = (byte)(128 + tmp % 128);
                        }
                        else
                        {
                            _buffX[localPosition++] = (byte)tmp;
                        }
                        _buffX[localPosition++] = (byte)(recLen >> 14);
                        if (_buffX[localPosition - 1] > 0)
                        {
                            _buffX[localPosition++] = (byte)0;
                        }
                    }
                    else
                    {
                        _buffX[localPosition++] = (byte)(recLen);
                        _buffX[localPosition++] = (byte)(recLen >> 8);
                    }


                    Int32ToSpecificBuffer(_buffX, (int)txtLen, localPosition);
                    localPosition += 4;
                    for (int i = 0; i < txtLen; i++)
                    {
                        _buffX[localPosition++] = (byte)(txt[i]); // = txt[i] % 256
                        _buffX[localPosition++] = (byte)(txt[i] >> 8); // = (byte)(txt[i]/256)

                        //_buff2[0] = (byte)(txt[i]); // = txt[i] % 256
                        //_buff2[1] = (byte)(txt[i] >> 8); // = (byte)(txt[i]/256)
                        //binaryWriter.Write(_buff2, 0, 2);
                        //binaryWriter.Write(BitConverter.GetBytes(txt[i]));
                    }
                    //binaryWriter.Write(_buffX, 0, localPosition);
                }
                binaryWriter.Write(_buffX, 0, localPosition);
                binaryWriter.Write(_endSst);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(_buffX);
            }
        }

        private static void RkNumberGeneralWrite(Span<byte> buff, double d/*, int offset, bool div100 = false*/)
        {
            // dla rk number
            // bytes[214] |=  0b00000001; = /100 flag
            // bytes[214] |=  0b00000010; = integer flag

            //if (div100)
            //{
            //    Int64 revD3 = BitConverter.DoubleToInt64Bits(100 * d);
            //    Int64 revD2 = revD3 >> 32;
            //    UInt32 revD1 = (uint)revD2;
            //    _buffer[offset + 0] = (byte)((revD1 % 256) | 0b00000001);
            //    revD1 >>= 8;
            //    _buffer[offset + 1] = (byte)(revD1 % 256);
            //    revD1 >>= 8;
            //    _buffer[offset + 2] = (byte)(revD1 % 256);
            //    revD1 >>= 8;
            //    _buffer[offset + 3] = (byte)(revD1 % 256);
            //}
            //else
            //{
            Int64 revD3 = BitConverter.DoubleToInt64Bits(d);
            Int64 revD2 = revD3 >> 32;
            UInt32 revD1 = (uint)revD2;

            buff[0] = (byte)((revD1 % 256) & 0b11111100);
            revD1 >>= 8;
            buff[1] = (byte)(revD1 % 256);
            revD1 >>= 8;
            buff[2] = (byte)(revD1 % 256);
            revD1 >>= 8;
            buff[3] = (byte)(revD1 % 256);
            //}
        }

        private static void Int32ToSpecificBuffer(Span<byte> _buff, Int32 intNumber, int offset)
        {
            BitConverter.TryWriteBytes(_buff[offset..], intNumber);
        }
    }
}
