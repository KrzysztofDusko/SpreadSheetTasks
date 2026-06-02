using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
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

        /// <summary>cellXfs index for general/numeric style.</summary>
        internal const int XfStyleGeneral = 0;
        /// <summary>cellXfs index for datetime style.</summary>
        internal const int XfStyleDateTime = 1;
        /// <summary>cellXfs index for date style.</summary>
        internal const int XfStyleDate = 2;
        /// <summary>cellXfs index for bold header (font 1).</summary>
        internal const int XfStyleBoldHeader = 3;

        /// <summary>
        /// Creates a new XLSB writer that will write to a file at <paramref name="path"/>.
        /// The file is overwritten if it already exists.
        /// </summary>
        /// <param name="path">Path of the .xlsb file to create.</param>
        /// <param name="compressionLevel">Compression level for the zip archive.</param>
        public XlsbWriter(string path, CompressionLevel compressionLevel = CompressionLevel.Fastest)
        {
            ArgumentNullException.ThrowIfNull(path);
            ArgumentException.ThrowIfNullOrWhiteSpace(path);

            FileStream fs = null!;
            try
            {
                fs = new FileStream(path, FileMode.Create);
                sheetCnt = 0;
                _compressionLevel = compressionLevel;
                _newExcelFileStream = fs;
                _excelArchiveFile = new ZipArchive(_newExcelFileStream, ZipArchiveMode.Create, false);
                _excelStreamWasProvided = false;
            }
            catch
            {
                fs?.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Creates a new XLSB writer that writes to the supplied <see cref="Stream"/>.
        /// </summary>
        /// <param name="stream">Destination stream. The writer takes ownership of the stream position but does not close it unless <paramref name="leaveExcelArchiveOpen"/> is false.</param>
        /// <param name="compressionLevel">Compression level for the zip archive.</param>
        /// <param name="leaveExcelArchiveOpen">If true, the underlying <see cref="ZipArchive"/> is left open when <see cref="Save"/> runs. Defaults to true for the stream ctor; the file-path ctor always sets it to false.</param>
        public XlsbWriter(Stream stream, CompressionLevel compressionLevel = CompressionLevel.Fastest, bool leaveExcelArchiveOpen = true)
        {
            ArgumentNullException.ThrowIfNull(stream);
            _excelStreamWasProvided = true;
            sheetCnt = 0;
            _compressionLevel = compressionLevel;
            _newExcelFileStream = stream;
            _excelArchiveFile = new ZipArchive(_newExcelFileStream, ZipArchiveMode.Create, leaveExcelArchiveOpen);
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

        public override void WriteSheet(IDataReader dataReader, Boolean headers = true, int maxRows = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            if (doAutofilter)
            {
                _autofilterIsOn = true;
                _filteredDict ??= new();
            }
            this._areHeaders = headers;
            _dataColReader = new DataColReader(dataReader, headers, maxRows);

            int rowNum = 0;
            _columnCount = _dataColReader.FieldCount;

            _startCol = (uint)startingColumn;
            _endCol = (uint)(_startCol + _columnCount);

            _colWidthsArray = new double[_columnCount];
            Array.Fill<double>(_colWidthsArray, -1.0);

            typesArray = new int[_columnCount];
            _newTypes = new TypeCode[_columnCount];

            var rdr = _dataColReader._dataReader ?? throw new InvalidOperationException("No IDataReader available. This overload requires an IDataReader source.");
            for (int l = 1; l <= _columnCount; l++)
            {
                int lenn = rdr.GetName(l - 1).Length + (doAutofilter?2:0);
                double tempWidth = 1.25 * lenn + 2;
                if (tempWidth > _MAX_WIDTH)
                {
                    tempWidth = _MAX_WIDTH;
                }
                if (_colWidthsArray[l - 1] < tempWidth)
                {
                    _colWidthsArray[l - 1] = tempWidth;
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
                SetColsLength(_columnCount, arr);
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
                InitSheet(doAutofilter);
                int writtenRowNum = 0;
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
                            ExcelWriter.SetTypes(_dataColReader, typesArray, _newTypes, _columnCount, detectBooleanType: true);
                        }
                    }

                    InitRow(writtenRowNum);
                    WriteRow(writtenRowNum == 0);
                    writtenRowNum++;

                    if (rowNum % 10000 == 0)
                    {
                        DoOn10k(rowNum);
                    }
                    rowNum++;
                }
                _rowsCount = writtenRowNum - 1;
                
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
            bool wroteAnyCell = false;
            for (int column = 0; column < _columnCount; column++)
            {
                if (_dataColReader.IsDBNull(column))
                    continue;

                wroteAnyCell = true;

                object rawVal = _dataColReader.GetValue(column);
                if (rawVal is FormattedCell fc)
                {
                    int fmtXfIndex = RegisterFormat(fc.Format);
                    WriteFormattedCellXlsb(fc.Value, fmtXfIndex, column);
                    continue;
                }

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
                    if (SuppressYear1000Dates && (dtVal as DateTime?).Value.Year == 1000)//1000-xx-xx
                    {
                        continue;
                    }
                    WriteDateTime(dtVal, column);
                 }
                
            }

            if (!wroteAnyCell)
            {
                WriteBlank(0, boldedStyle ? 3 : 0);
            }
        }

        private void WriteFormattedCellXlsb(object val, int xfIndex, int column)
        {
            if (val is string s)
            {
                WriteString(s, column, bolded: false, styleOverride: xfIndex);
            }
            else if (val is bool b)
            {
                WriteBool(b, column, xfIndex);
            }
            else if (val is DateTime dt)
            {
                WriteDouble(dt.ToOADate(), column, xfIndex);
            }
            else if (val is sbyte sb) { WriteDouble(sb, column, xfIndex); }
            else if (val is byte ub) { WriteDouble(ub, column, xfIndex); }
            else if (val is short sVal) { WriteDouble(sVal, column, xfIndex); }
            else if (val is ushort usVal) { WriteDouble(usVal, column, xfIndex); }
            else if (val is int iVal)
            {
                if (iVal >= _rRkIntegerLowerLimit && iVal <= _rRkIntegerUpperLimit)
                    WriteRkNumberInteger(iVal, column, xfIndex);
                else
                    WriteDouble((double)iVal, column, xfIndex);
            }
            else if (val is uint uiVal) { WriteDouble(uiVal, column, xfIndex); }
            else if (val is long lVal) { WriteDouble(lVal, column, xfIndex); }
            else if (val is ulong ulVal) { WriteDouble(ulVal, column, xfIndex); }
            else if (val is float fVal) { WriteDouble(fVal, column, xfIndex); }
            else if (val is double dVal) { WriteDouble(dVal, column, xfIndex); }
            else if (val is decimal mVal) { WriteDouble((double)mVal, column, xfIndex); }
            else
            {
                WriteString(val?.ToString() ?? "", column, bolded: false, styleOverride: xfIndex);
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
                _stream.WriteByte((byte)(_colWidthsArray[l])); // .. x 7 = pixels
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

        private void WriteDouble(double val, int colNum, int styleNum = 0)
        {
            Span<byte> buff = stackalloc byte[18];
            buff[0] = 5;
            buff[1] = 16;
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);
            BitConverter.TryWriteBytes(buff[6..], styleNum);
            BitConverter.TryWriteBytes(buff[10..], val);
            _stream.Write(buff);
        }

        private void WriteBool(bool val, int colNum, int styleNum = 0)
        {
            Span<byte> buff = stackalloc byte[11];
            buff[0] = 0x04;
            buff[1] = 8 + 1;
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);
            BitConverter.TryWriteBytes(buff[6..], styleNum);
            buff[10] = (byte)(val ? 1 : 0);
            _stream.Write(buff);
        }

        private void WriteBlank(int colNum, int styleNum = 0)
        {
            Span<byte> buff = stackalloc byte[10];
            buff[0] = 0x01;
            buff[1] = 8;
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);
            BitConverter.TryWriteBytes(buff[6..], styleNum);
            _stream.Write(buff);
        }

        private void WriteRkNumberInteger(int val, int colNum, int styleNum = 0)
        {
            Span<byte> buff = stackalloc byte[14];
            buff[0] = 2;
            buff[1] = 12;
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);
            BitConverter.TryWriteBytes(buff[6..], styleNum);

            val <<= 2;
            val |= 0b00000010;

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


        private void WriteString(string stringValue, int colNum, bool bolded = false, int styleOverride = -1)
        {
            ref var index = ref CollectionsMarshal.GetValueRefOrAddDefault(_sstDic, stringValue, out bool exists);
            if (!exists)
            {
                index = _sstCntUnique;
                _sstCntUnique++;
            }
            WriteStringFromShared(index, colNum, bolded, styleOverride);
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
                buff[1] = 12;
                BitConverter.TryWriteBytes(buff[2..], (int)colNum);
                BitConverter.TryWriteBytes(buff[6..], 2);
                RkNumberGeneralWrite(buff[10..],d1);
                _stream.Write(buff);
            }
        }

        private void WriteStringFromShared(int val, int colNum, bool bolded = false, int styleOverride = -1)
        {
            Span<byte> buff = stackalloc byte[14];
            buff[0] = 7;
            buff[1] = 12;
            BitConverter.TryWriteBytes(buff[2..], (int)colNum);

            int styleNum = styleOverride >= 0 ? styleOverride : (bolded ? 3 : 0);
            BitConverter.TryWriteBytes(buff[6..], styleNum);

            BitConverter.TryWriteBytes(buff[10..], (int)val);
            _stream.Write(buff);
        }

        private static void WriteVarint(Stream s, uint value)
        {
            while (value >= 0x80)
            {
                s.WriteByte((byte)((value & 0x7F) | 0x80));
                value >>= 7;
            }
            s.WriteByte((byte)value);
        }

        private static void WriteBrtRecord(Stream s, uint type, byte[] data)
        {
            WriteVarint(s, type);
            WriteVarint(s, (uint)data.Length);
            s.Write(data);
        }

        private static byte[] BuildStFmt(int ifmt, string fmtString)
        {
            int cch = fmtString.Length;
            int remaining = 2 + 4 + cch * 2;
            var data = new byte[2 + remaining];
            data[0] = 0x2C;
            data[1] = (byte)remaining;
            BitConverter.TryWriteBytes(data.AsSpan(2, 2), (ushort)ifmt);
            BitConverter.TryWriteBytes(data.AsSpan(4, 4), cch);
            Encoding.Unicode.GetBytes(fmtString, 0, cch, data, 8);
            return data;
        }

        private static byte[] BuildXfRecord(int fontId, int ifmt, int flags, int byte6Extra = 0)
        {
            var data = new byte[18];
            data[0] = 0x2F;
            data[1] = 16;
            BitConverter.TryWriteBytes(data.AsSpan(2, 2), (ushort)fontId);
            BitConverter.TryWriteBytes(data.AsSpan(4, 2), (ushort)ifmt);
            data[6] = (byte)byte6Extra;
            data[14] = 0x10;
            data[15] = 0x10;
            BitConverter.TryWriteBytes(data.AsSpan(16, 2), (ushort)flags);
            return data;
        }

        private byte[] BuildStylesBytes()
        {
            if (!_hasCustomFormats)
                return _stylesBin;

            using var ms = new MemoryStream();

            // BrtBeginStyleSheet
            ms.Write(_stylesBin, 0, 3);

            // ── BrtFmt (type 0x0267): header with count ──
            var stFmtList = new List<byte[]>();
            stFmtList.Add(BuildStFmt(164, "yyyy\\-mm\\-dd\\ hh:mm"));
            stFmtList.Add(BuildStFmt(166, "yyyy\\-mm\\-dd"));
            foreach (var kvp in _formatRegistry)
                stFmtList.Add(BuildStFmt(kvp.Value, kvp.Key));

            int fmtCount = 2 + _formatRegistry.Count;
            var fmtHeader = new byte[4];
            BitConverter.TryWriteBytes(fmtHeader, (ushort)fmtCount);
            WriteBrtRecord(ms, 0x0267, fmtHeader);

            // stFmt records as separate Brt sub-records (type=0x2C)
            foreach (var rec in stFmtList)
                ms.Write(rec);

            // BrtEndFonts (type 0x0268)
            WriteBrtRecord(ms, 0x0268, Array.Empty<byte>());

            // ── Copy Fills, Borders, CellStyleXFs from template ──
            int fillStart = Array.IndexOf(_stylesBin, (byte)0xE3);
            if (fillStart < 0) return _stylesBin;

            int cellXfRecStart = -1;
            for (int i = fillStart; i < _stylesBin.Length - 1; i++)
            {
                if (_stylesBin[i] == 0xE9 && _stylesBin[i + 1] == 0x04)
                { cellXfRecStart = i; break; }
            }
            if (cellXfRecStart < fillStart) return _stylesBin;

            int cellXfRecEnd = -1;
            for (int i = cellXfRecStart + 1; i < _stylesBin.Length - 2; i++)
            {
                if (_stylesBin[i] == 0xEA && _stylesBin[i + 1] == 0x04 && _stylesBin[i + 2] == 0x00)
                { cellXfRecEnd = i; break; }
            }
            if (cellXfRecEnd < cellXfRecStart) return _stylesBin;

            // Find and copy through BrtEndCellStyleXFs (before cellXFs section)
            // BrtBeginCellStyleXFs (0x0272) → [0xF2,0x04]
            int cellStyleXfEnd = -1;
            for (int i = fillStart; i < cellXfRecStart - 2; i++)
            {
                if (_stylesBin[i] == 0xF3 && _stylesBin[i + 1] == 0x04 && _stylesBin[i + 2] == 0x00)
                { cellStyleXfEnd = i; break; }
            }

            if (cellStyleXfEnd > fillStart)
                ms.Write(_stylesBin, fillStart, cellStyleXfEnd + 3 - fillStart);
            else
                ms.Write(_stylesBin, fillStart, cellXfRecStart - fillStart);

            // ── Build custom cellXFs section ──
            int totalXf = 4 + _formatXfMap.Count;

            var xfHeader = new byte[4];
            BitConverter.TryWriteBytes(xfHeader, (ushort)totalXf);
            WriteBrtRecord(ms, 0x0269, xfHeader);

            ms.Write(BuildXfRecord(0, 0, 0x0000));
            ms.Write(BuildXfRecord(0, 164, 0x0001));
            ms.Write(BuildXfRecord(0, 166, 0x0001));
            ms.Write(BuildXfRecord(1, 0, 0x0000, 1));

            foreach (var kvp in _formatXfMap.OrderBy(k => k.Value))
            {
                int numFmtId = _formatRegistry[kvp.Key];
                ms.Write(BuildXfRecord(0, numFmtId, 0x0001));
            }

            WriteBrtRecord(ms, 0x026A, Array.Empty<byte>());

            // ── Copy remaining template ──
            int afterCellXf = cellXfRecEnd + 3;
            ms.Write(_stylesBin, afterCellXf, _stylesBin.Length - afterCellXf);

            return ms.ToArray();
        }

        internal override void FinalizeFile()
        {
            SaveSst();
            var newEntry = _excelArchiveFile.CreateEntry(@"xl/styles.bin", _compressionLevel);
            using (var str = newEntry.Open())
            {
                using var sw = new BinaryWriter(str);
                sw.Write(BuildStylesBytes());
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

                if (!String.IsNullOrWhiteSpace(DocPropertyProgramName))
                {
                    sw.Write("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
                    sw.Write("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
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
                if (!String.IsNullOrWhiteSpace(DocPropertyProgramName))
                {
                    sw.Write("<Relationship Id=\"rId2\" ");
                    sw.Write("Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" ");
                    sw.Write("Target=\"docProps/core.xml\"/>");
                    sw.Write("<Relationship Id=\"rId3\" ");
                    sw.Write("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" ");
                    sw.Write("Target=\"docProps/app.xml\"/>");
                }
                sw.Write(@"<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.bin""/>");
                sw.Write(@"</Relationships>");
            }

            if (!String.IsNullOrWhiteSpace(DocPropertyProgramName))
            {
                var e2 = _excelArchiveFile.CreateEntry("docProps/app.xml", _compressionLevel);
                using (var str = e2.Open())
                using (var sw = new StreamWriter(str, Encoding.UTF8))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties ");
                    sw.Write("xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" ");
                    sw.Write("xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
                    sw.Write($"<Application>{DocPropertyProgramName}</Application>");
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
                    using var coreSw = new StreamWriter(str, Encoding.UTF8);
                    coreSw.WriteLine($@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                    coreSw.Write($@"<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">");
                    coreSw.Write($@"<dc:creator>{DocPropertyProgramName} - used by {Environment.UserName}</dc:creator>");
                    coreSw.Write($@"<cp:lastModifiedBy>{DocPropertyProgramName} - used by {Environment.UserName}</cp:lastModifiedBy>");
                    coreSw.Write($@"<dcterms:created xsi:type=""dcterms:W3CDTF"">2015-06-05T18:19:34Z</dcterms:created>");
                    coreSw.Write($@"<dcterms:modified xsi:type=""dcterms:W3CDTF"">2021-09-05T11:11:46Z</dcterms:modified>");
                    coreSw.Write($@"</cp:coreProperties>");
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
