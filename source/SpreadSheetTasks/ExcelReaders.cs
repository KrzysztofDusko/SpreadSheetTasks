using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace SpreadSheetTasks
{
    public sealed class XlsxOrXlsbReadOrEdit : ExcelReaderAbstract, IDisposable
    {
        private ZipArchive _xlsxArchive;
        private readonly Dictionary<string, string> _worksheetIdToName = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _worksheetNameToId = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _worksheetIdToLocation = new Dictionary<string, string>();
        private Dictionary<int, string> _pivotCacheIdtoRid;
        private Dictionary<string, string> _pivotCachRidToLocation;
        //private Dictionary<int, string> _worksheetSheetIdToId = new Dictionary<int, string>();


        private string[] _sharedStringArray;
        private StyleInfo[] _stylesCellXfsArray;

        internal readonly static Dictionary<int, Type> _numberFormatsTypeDic = new Dictionary<int, Type>()
        {
            // to do
            {0,typeof(string)},
            {1,typeof(decimal?)},
            {2,typeof(decimal?)},
            {3,typeof(decimal?)},
            {4,typeof(decimal?)},
            {5,typeof(decimal?)},
            {6,typeof(decimal?)},
            {7,typeof(decimal?)},

            {9, typeof(decimal?)}, // 0%

            { 14,typeof(DateTime?)},
            { 15,typeof(DateTime?)},
            { 16,typeof(DateTime?)},
            { 17,typeof(DateTime?)},
            { 18,typeof(DateTime?)},
            { 19,typeof(DateTime?)},
            { 20,typeof(DateTime?)},
            { 21,typeof(DateTime?)},
            { 22,typeof(DateTime?)},

            { 44,typeof(decimal?)}
        };

        internal readonly static HashSet<string> _dateExcelMasks = new HashSet<string>()
        {
            @"[$-F800]dddd\,\ mmmm\ dd\,\ yyyy",
            @"d\-mm;@",
            @"yy\-mm\-dd;@",
            @"[$-415]d\ mmm;@",
            @"[$-415]d\ mmm\ yy;@",
            @"[$-415]dd\ mmm\ yy;@",
            @"[$-415]mmm\ yy;@",
            @"[$-415]mmmm\ yy;@",
            @"[$-415]d\ mmmm\ yyyy;@",
            @"yyyy\-mm\-dd\ hh:mm",
            @"yyyy\-mm\-dd\ hh:mm:ss"
        };

        private string _sharedStringsLocation = null;
        private string _stylesLocation = null;
        private string _themeLocation = null;
        private int _uniqueStringCount = -1;
        private int _stringCount = -1;
        private static readonly CultureInfo invariantCultureInfo = CultureInfo.InvariantCulture;

        Modes mode = Modes.xlsx;
        enum Modes
        {
            xlsx, xlsb
        }

        private static readonly string _openXmlInfoString = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private static readonly XmlReaderSettings _xmlSettings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            XmlResolver = null
        };

        private static readonly Dictionary<string, int> _letterToColumnNum = new Dictionary<string, int>();
        static XlsxOrXlsbReadOrEdit()
        {
            for (int i = 0; i < XlsxWriter._letters.Length; i++)
            {
                _letterToColumnNum[XlsxWriter._letters[i]] = i;
            }
        }

        public override void Dispose()
        {
            if (_xlsxArchive != null)
            {
                _xlsxArchive.Dispose();
            }
        }

        private void OpenXlsx(string path, bool readSharedStrings = true, bool updateMode = false)
        {
            mode = Modes.xlsx;
            if (updateMode)
            {
                _xlsxArchive = new ZipArchive(new FileStream(path, FileMode.Open), ZipArchiveMode.Read | ZipArchiveMode.Update);
            }
            else
            {
                _xlsxArchive = new ZipArchive(new FileStream(path, FileMode.Open), ZipArchiveMode.Read);
            }

            var e1 = _xlsxArchive.GetEntry("xl/workbook.xml");
            using (var str = e1.Open())
            {
                using var reader = XmlReader.Create(str);
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Whitespace)
                    {
                        continue;
                    }

                    if (reader.Name == "sheet")
                    {
                        string name = reader.GetAttribute("name");
                        string rId = reader.GetAttribute("r:id");
                        _worksheetIdToName[rId] = name;
                        _worksheetNameToId[name] = rId;

                        //int.TryParse(reader.GetAttribute("sheetId"), out int sheetId);
                        //_worksheetSheetIdToId[sheetId] = rId;
                    }
                    else if (reader.Name == "pivotCache")
                    {
                        if (_pivotCacheIdtoRid == null)
                        {
                            _pivotCacheIdtoRid = new Dictionary<int, string>();
                        }

                        if (!int.TryParse(reader.GetAttribute("cacheId"), out int cacheId))
                        {
                            throw new Exception("getting pivot table cache error");
                        }

                        string rId = reader.GetAttribute("r:id");
                        _pivotCacheIdtoRid[cacheId] = rId;
                    }
                }
            }

            _resultCount = _worksheetIdToName.Count;
            FillRels("xml");

            if (_stylesLocation != null)
            {
                FillStyles();
            }
            if (readSharedStrings && _sharedStringsLocation != null)
            {
                FillSharedStrings();
            }
        }

        private void OpenXlsb(string path, bool readSharedStrings = true)
        {
            mode = Modes.xlsb;
            _xlsxArchive = new ZipArchive(new FileStream(path, FileMode.Open), ZipArchiveMode.Read);

            var e1 = _xlsxArchive.GetEntry("xl/workbook.bin");

            Stream str;
            if (UseMemoryStreamInXlsb)
            {
                str = GetMemoryStream(e1.Open(), e1.Length);
            }
            else
            {
                str = new BufferedStream(e1.Open());
            }

            try
            {
                using var reader = new BiffReaderWriter(str);
                while (reader.ReadWorkbook())
                {
                    if (reader.isSheet == true)
                    {
                        string name = reader.workbookName;
                        string rId = reader.recId;
                        _worksheetIdToName[rId] = name;
                        _worksheetNameToId[name] = rId;
                    }
                }
            }
            finally
            {
                str.Dispose();
            }


            _resultCount = _worksheetIdToName.Count;

            FillRels("bin");


            if (_stylesLocation != null)
            {
                FillBinStyles();
            }

            if (readSharedStrings && _sharedStringsLocation != null)
            {
                FillBinSharedStrings();
            }

        }

        public override void Open(string path, bool readSharedStrings = true, bool updateMode = false)
        {
            if (path.EndsWith("xlsb", StringComparison.OrdinalIgnoreCase))
            {
                mode = Modes.xlsb;
                OpenXlsb(path, readSharedStrings);
            }
            else
            {
                mode = Modes.xlsx;
                OpenXlsx(path);
            }
        }

        private void FillRels(string xmlOrBin)
        {
            var e2 = _xlsxArchive.GetEntry($"xl/_rels/workbook.{xmlOrBin}.rels");
            using var str = e2.Open();
            using var reader = XmlReader.Create(str);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Whitespace)
                {
                    continue;
                }

                if (reader.Name == "Relationship")
                {
                    string target = reader.GetAttribute("Target");
                    string type = reader.GetAttribute("Type");
                    string rId = reader.GetAttribute("Id");

                    if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                    {
                        _worksheetIdToLocation[rId] = target;
                    }
                    else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition")
                    {
                        if (_pivotCachRidToLocation == null)
                        {
                            _pivotCachRidToLocation = new Dictionary<string, string>();
                        }
                        _pivotCachRidToLocation[rId] = target;
                    }
                    else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
                    {
                        _sharedStringsLocation = target;
                    }
                    else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
                    {
                        _stylesLocation = target;
                    }
                    else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
                    {
                        _themeLocation = target;
                    }

                }
            }
        }

        public bool UseMemoryStreamInXlsb = true;
        private const int bufferSize = 256 * 256;
        public static MemoryStream GetMemoryStream(Stream streamToRead, long length)
        {
            byte[] byteArray = new byte[length];
            int pos = 0;
            int bytesRead = 0;
            int toRead = bufferSize;
            while (true)
            {
                if (length - bytesRead < toRead)
                {
                    toRead = (int)length - bytesRead;
                }

                pos = streamToRead.Read(byteArray, bytesRead, toRead);
                bytesRead += pos; // pos = 4096 in most cases
                if (pos == 0)
                {
                    break;
                }
            }
            return new MemoryStream(byteArray);
        }

        public void FillSharedStrings()
        {
            if (_sharedStringsLocation == null)
            {
                throw new Exception("no shared strings found");
            }

            var _sharedStringList = new List<string>();
            var sharedstringsEntry = _xlsxArchive.GetEntry(@$"xl/{_sharedStringsLocation}");

            Stream str = sharedstringsEntry.Open();
            //Stream str = getMemoryStream(sharedstringsEntry.Open(), sharedstringsEntry.Length);

            try
            {
                using var reader = XmlReader.Create(str/*, _xmlSettings*/);
                reader.Read();
                if (reader.IsStartElement("sst", _openXmlInfoString))
                {
                    string unqCnt = reader.GetAttribute("uniqueCount");
                    if (unqCnt != null)
                    {
                        int.TryParse(unqCnt, out _uniqueStringCount);
                    }

                    string cnt = reader.GetAttribute("count");
                    if (cnt != null)
                    {
                        int.TryParse(cnt, out _stringCount);
                    }
                }
                else
                {
                    throw new Exception("not openXml 2006 format!");
                }

                while (reader.Read())
                {
                    if (reader.IsStartElement("si"/*, _openXmlInfoString*/))
                    {
                        reader.Read();
                        if (reader.IsStartElement("t"/*, _openXmlInfoString*/))
                        {
                            string vall = reader.ReadElementContentAsString();
                            // si -> t-> wnętrze t -> wartość
                            //string vall = reader.Value;
                            _sharedStringList.Add(vall);
                        }
                    }
                }
            }
            finally
            {
                str.Dispose();
            }

            _sharedStringArray = _sharedStringList.ToArray();
        }

        public void FillBinSharedStrings()
        {
            if (_sharedStringsLocation == null)
            {
                throw new Exception("no shared strings found");
            }

            var _sharedStringList = new List<string>();
            var sharedstringsEntry = _xlsxArchive.GetEntry(@$"xl/{_sharedStringsLocation}");


            Stream str;
            if (UseMemoryStreamInXlsb)
            {
                str = GetMemoryStream(sharedstringsEntry.Open(), sharedstringsEntry.Length);
            }
            else
            {
                str = new BufferedStream(sharedstringsEntry.Open());
            }

            try
            {
                using var reader = new BiffReaderWriter(str);
                while (reader.ReadSharedStrings())
                {
                    string vall = reader.sharedStringValue;
                    if (vall != null)
                    {
                        _sharedStringList.Add(vall);
                    }
                }
            }
            finally
            {
                str.Dispose();
            }

            _sharedStringArray = _sharedStringList.ToArray();
        }

        private void FillStyles()
        {
            var _stylesCellXfs = new List<StyleInfo>();
            //var _styleCellStyleXfs = new List<StyleInfo>();
            //_customNumberFormatsDic = new Dictionary<int, NumberFormatInfo>();

            var e = _xlsxArchive.GetEntry($"xl/{_stylesLocation}");
            using (var str = e.Open())
            {
                using var reader = XmlReader.Create(str);
                reader.Read();
                while (!reader.EOF)
                {
                    if (reader.IsStartElement("cellXfs"))
                    {
                        reader.Read();
                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("xf"))
                            {
                                int.TryParse(reader.GetAttribute("xfId"), out var xfId);
                                int.TryParse(reader.GetAttribute("numFmtId"), out var numFmtId);

                                /*if (!int.TryParse(reader.GetAttribute("applyNumberFormat"), out var applyNumberFormat))
                                {
                                    applyNumberFormat = -1;
                                }*/

                                _stylesCellXfs.Add(new StyleInfo() { XfId = xfId, NumFmtId = numFmtId/*, ApplyNumberFormat = applyNumberFormat*/ });
                                reader.Skip();
                            }
                            else
                            {
                                break;
                            }
                        }
                    }

                    //else if (reader.IsStartElement("cellStyleXfs"))
                    //{
                    //    reader.Read();
                    //    while (!reader.EOF)
                    //    {
                    //        if (reader.IsStartElement("xf"))
                    //        {
                    //            int.TryParse(reader.GetAttribute("xfId"), out var xfId);
                    //            int.TryParse(reader.GetAttribute("numFmtId"), out var numFmtId);

                    //            if (!int.TryParse(reader.GetAttribute("applyNumberFormat"), out var applyNumberFormat))
                    //            {
                    //                applyNumberFormat = -1;
                    //            }
                    //            _styleCellStyleXfs.Add(new StyleInfo() { XfId = xfId, NumFmtId = numFmtId, ApplyNumberFormat = applyNumberFormat });
                    //            reader.Skip();
                    //        }
                    //        else
                    //        {
                    //            break;
                    //        }
                    //    }
                    //}

                    else if (reader.IsStartElement("numFmts"))
                    {
                        reader.Read();
                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("numFmt"))
                            {
                                string formatCode = reader.GetAttribute("formatCode");
                                int.TryParse(reader.GetAttribute("numFmtId"), out var numFmtId);

                                if (!_numberFormatsTypeDic.TryGetValue(numFmtId, out var type))
                                {
                                    if (_dateExcelMasks.Contains(formatCode))
                                    {
                                        type = typeof(DateTime?);
                                    }
                                    else
                                    {
                                        type = typeof(string);
                                    }
                                    _numberFormatsTypeDic[numFmtId] = type;
                                }
                                else
                                {
                                    _numberFormatsTypeDic[numFmtId] = type;
                                }
                                //_customNumberFormatsDic[numFmtId] = new NumberFormatInfo
                                //{
                                //    Name = formatCode,
                                //    proposedType = type
                                //};

                                //if (!_numberFormatsTypeDic.ContainsKey(numFmtId))
                                //{
                                //    _numberFormatsTypeDic[numFmtId] = type;
                                //}

                                reader.Skip();
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    if (reader.Depth > 0)
                    {
                        reader.Skip();
                    }
                    else
                    {
                        reader.Read();
                    }
                }
            }
            _stylesCellXfsArray = _stylesCellXfs.ToArray();
            //_stylesCellStyleXfsArray = _styleCellStyleXfs.ToArray();
        }

        private void FillBinStyles()
        {
            var _stylesCellXfs = new List<StyleInfo>();
            //var _styleCellStyleXfs = new List<StyleInfo>();
            //_customNumberFormatsDic = new Dictionary<int, NumberFormatInfo>();

            var e = _xlsxArchive.GetEntry($"xl/{_stylesLocation}");
            using (var str = e.Open())
            {
                bool stylesFirstTime = true;
                bool formatFirstTime = true;

                using var reader = new BiffReaderWriter(str);
                while (reader.ReadStyles())
                {
                    if (reader._inCellXf)
                    {
                        if (stylesFirstTime)
                        {
                            reader.ReadStyles();
                            stylesFirstTime = false;
                        }
                        int numFmtId = reader.NumberFormatIndex;
                        int xfId = reader.ParentCellStyleXf;
                        _stylesCellXfs.Add(new StyleInfo() { XfId = xfId, NumFmtId = numFmtId });
                        //reader.ReadStyles();
                    }
                    else if (reader._inNumberFormat)
                    {
                        if (formatFirstTime)
                        {
                            reader.ReadStyles();
                            formatFirstTime = false;
                        }

                        string formatCode = reader.formatString;
                        int numFmtId = reader.format;

                        if (!_numberFormatsTypeDic.TryGetValue(numFmtId, out var type))
                        {
                            if (_dateExcelMasks.Contains(formatCode))
                            {
                                type = typeof(DateTime?);
                            }
                            else
                            {
                                type = typeof(string);
                            }
                            _numberFormatsTypeDic[numFmtId] = type;
                        }
                        //reader.ReadStyles();
                    }
                }
            }
            _stylesCellXfsArray = _stylesCellXfs.ToArray();
            //_stylesCellStyleXfsArray = _styleCellStyleXfs.ToArray();
        }

        public override string[] GetScheetNames()
        {
            return _worksheetIdToName.Values.ToArray();
        }

        private static readonly char[] _digits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

        private string _actualSheetDimensions = null;

        private object[] _blankRow;
        private readonly char[] _buffer = new char[64];
        public IEnumerable<object[]> GetRowsOfXlsx(string sheetName)
        {
            bool firstRowX = true;
            int columnsCntFromFirstRowX = -1;
            int minColNumX = -1;
            int maxColNum = -1;
            int collDif = 0;

            ZipArchiveEntry sheetEntry = GetArchiverEntry(sheetName);
            Stream str = sheetEntry.Open();

            try
            {
                using var reader = XmlReader.Create(str, _xmlSettings);
                reader.Read();
                while (!reader.EOF)
                {
                    if (reader.IsStartElement("sheetData"))
                    {
                        int prevRow = -1;

                        while (reader.Read() && reader.IsStartElement("row"))
                        {
                            reader.MoveToAttribute("r");
                            int len = reader.ReadValueChunk(_buffer, 0, _buffer.Length);
                            int rowNum = ParseToUnsignedIntFromBuffer(_buffer, len);

                            //empty row/s
                            if (prevRow != -1 && rowNum > prevRow + 1)
                            {
                                if (_blankRow == null)
                                {
                                    _blankRow = new object[columnsCntFromFirstRowX];
                                }

                                for (int i = 0; i < rowNum - prevRow - 1; i++)
                                {
                                    yield return _blankRow;
                                }
                            }
                            //reader.GetAttribute("spans")
                            if (firstRowX)
                            {
                                innerRow = ArrayPool<object>.Shared.Rent(128 * 128);
                            }

                            prevRow = rowNum;

                            int colNum = -1;
                            int prevColNum = -1;
                            while (reader.Read() && reader.IsStartElement("c"))
                            {
                                reader.MoveToAttribute("r");
                                len = reader.ReadValueChunk(_buffer, 0, _buffer.Length);
                                colNum = -1;
                                char c;
                                for (int j = 0; j < len; j++)
                                {
                                    c = _buffer[j];
                                    if (c < 'A' || c > 'Z')
                                    {
                                        break;
                                    }
                                    int v = _buffer[j] - 'A';
                                    if ((uint)v < 26u)
                                    {
                                        colNum = ((colNum + 1) * 26) + v;
                                    }
                                }
                                colNum++;

                                if (firstRowX && minColNumX == -1)
                                {
                                    minColNumX = colNum;
                                    collDif = minColNumX - 1;
                                }
                                if (firstRowX && maxColNum < colNum)
                                {
                                    maxColNum = colNum;
                                }

                                object valueX = null;
                                bool isEmptyElement = reader.IsEmptyElement;
                                if (!isEmptyElement)
                                {
                                    char sstMark = '\0';
                                    int sstLen = 0;
                                    if (reader.MoveToAttribute("t"))
                                    {
                                        sstLen = reader.ReadValueChunk(_buffer, 0, _buffer.Length);
                                        sstMark = _buffer[0];
                                    }

                                    bool isStyle = reader.MoveToAttribute("s");
                                    int styleId = -1;
                                    if (isStyle)
                                    {
                                        len = reader.ReadValueChunk(_buffer, 0, _buffer.Length);
                                        styleId = ParseToUnsignedIntFromBuffer(_buffer, len);
                                    }

                                    //if (reader.IsStartElement("f")) // skip functions !
                                    //{
                                    //    reader.Skip();
                                    //}
                                    //reader.Read();
                                    reader.MoveToElement();
                                    bool success = false;
                                    if (sstMark == 'i' && sstLen == 9)
                                    {
                                        success = reader.ReadToDescendant("is");
                                        reader.Read();
                                        reader.Read();
                                        valueX = reader.ReadContentAsString();
                                    }
                                    else if (reader.ReadToDescendant("v"))
                                    {
                                        reader.Read();

                                        if (sstMark == 's' && sstLen == 1) // 's' = string/sharedstring'b' = boolean, 'e' = error, shared strings
                                        {
                                            len = reader.ReadValueChunk(_buffer, 0, _buffer.Length);
                                            valueX = _sharedStringArray[ParseToUnsignedIntFromBuffer(_buffer, len)];
                                        }
                                        else if (sstMark == 's' || sstMark == 'i' && sstLen == 9)  // InlineStr?
                                        {
                                            reader.Read();
                                            valueX = reader.ReadContentAsString();
                                            //valueX = _buffer.AsSpan(0, len).ToString();
                                        }
                                        else
                                        {
                                            len = reader.ReadValueChunk(_buffer, 0, _buffer.Length);
                                            if (sstMark == 'b') // 'b' = boolean, 'e' = error
                                            {
                                                valueX = (_buffer[0] == '1');
                                            }
                                            else if (sstMark == 'e') // 'b' = boolean, 'e' = error
                                            {
                                                valueX = "error in cell";
                                            }
                                            else if (styleId != -1)
                                            {
                                                var s = _stylesCellXfsArray[styleId];
                                                int numFormatId = s.NumFmtId;

                                                if (_numberFormatsTypeDic.TryGetValue(numFormatId, out Type type) && type == typeof(DateTime?)
                                                    && double.TryParse(_buffer.AsSpan(0, len), NumberStyles.Any, provider: invariantCultureInfo, out double doubleDate)
                                                    )
                                                {
                                                    valueX = DateTime.FromOADate(doubleDate);
                                                }
                                                else
                                                {
                                                    if (ContainsDoubleMarks(_buffer, len))
                                                    {
                                                        valueX = double.Parse(_buffer.AsSpan(0, len), provider: invariantCultureInfo);
                                                    }
                                                    else
                                                    {
                                                        valueX = ParseToInt64FromBuffer(_buffer, len);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (ContainsDoubleMarks(_buffer, len))
                                                {
                                                    valueX = double.Parse(_buffer.AsSpan(0, len), provider: invariantCultureInfo);
                                                }
                                                else
                                                {
                                                    valueX = ParseToInt64FromBuffer(_buffer, len);
                                                }
                                            }
                                        }
                                    }
                                }

                                if (colNum == prevColNum + 1 || prevColNum == -1 && colNum == minColNumX)
                                {
                                    innerRow[colNum - 1 - collDif] = valueX;
                                }
                                else if (prevColNum == -1 && colNum != minColNumX)
                                {
                                    for (int i = minColNumX; i < colNum; i++)
                                    {
                                        innerRow[i - 1 - collDif] = null;
                                    }
                                    innerRow[colNum - 1 - collDif] = valueX;
                                }
                                else if (colNum - prevColNum > 1)
                                {
                                    for (int i = 0; i < colNum - prevColNum - collDif; i++)
                                    {
                                        innerRow[prevColNum - collDif + i] = null;
                                    }
                                    innerRow[colNum - 1 - collDif] = valueX;
                                }
                                prevColNum = colNum;

                                if (!isEmptyElement) // depth = ...
                                {
                                    while (reader.Depth > 3)
                                    {
                                        reader.Skip();
                                    }
                                }
                            }

                            if (firstRowX)
                            {
                                firstRowX = false;
                                columnsCntFromFirstRowX = maxColNum - minColNumX + 1;
                                var arr = new object[columnsCntFromFirstRowX];
                                Array.Copy(innerRow, arr, columnsCntFromFirstRowX);
                                ArrayPool<object>.Shared.Return(innerRow);
                                innerRow = arr;

                                FieldCount = columnsCntFromFirstRowX;
                            }
                            if (colNum < maxColNum)
                            {
                                for (int i = colNum; i < maxColNum; i++)
                                {
                                    innerRow[i - collDif] = null;
                                }
                            }
                            yield return innerRow;
                        }
                    }
                    else if (reader.IsStartElement("dimension"))
                    {
                        _actualSheetDimensions = reader.GetAttribute("ref");
                        reader.Skip();
                    }
                    else if (reader.Depth == 0)
                    {
                        reader.Read();
                    }
                    else
                    {
                        reader.Skip();
                    }
                }
                yield break;
            }
            finally
            {
                str.Dispose();
            }
        }

        static int ParseToUnsignedIntFromBuffer(char[] buff, int len)
        {
            int res = 0;
            for (int i = 0; i < len; i++)
            {
                res = res * 10 + (buff[i] - '0');
            }
            return res;
        }

        static Int64 ParseToInt64FromBuffer(char[] buff, int len)
        {
            Int64 res = 0;
            int start = buff[0] == '-' ? 1 : 0;
            for (int i = start; i < len; i++)
            {
                res = res * 10 + (buff[i] - '0');
            }
            return start == 1 ? -res : res;
        }

        static bool ContainsDoubleMarks(char[] buff, int len)
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

        public IEnumerable<object[]> GetRowsOfXlsb(string sheetName)
        {
            bool firstRowX = true;
            int columnsCntFromFirstRowX = -1;
            int minColNumX = -1;
            int maxColNum = -1;
            int colNum = -1;

            ZipArchiveEntry sheetEntry = GetArchiverEntry(sheetName);

            Stream str;

            if (UseMemoryStreamInXlsb)
            {
                str = GetMemoryStream(sheetEntry.Open(), sheetEntry.Length);
            }
            else
            {
                str = new BufferedStream(sheetEntry.Open());
            }
            int column = 0;
            innerRow = new object[1024];
            try
            {
                using var reader = new BiffReaderWriter(str);
                int prevRow = -1;
                int prevColNum = -1;
                while (reader.ReadWorksheet())
                {
                    if (reader.readCell)
                    {
                        int rowNum = reader.rowIndex;

                        if (rowNum != prevRow && prevRow != -1)
                        {
                            if (firstRowX)
                            {
                                firstRowX = false;
                                columnsCntFromFirstRowX = column;
                                FieldCount = columnsCntFromFirstRowX;
                                var _rowListVlues2 = new object[FieldCount];
                                Array.Copy(innerRow, _rowListVlues2, FieldCount);
                                innerRow = _rowListVlues2;
                            }
                            if (colNum < maxColNum)
                            {
                                for (int i = colNum; i < maxColNum; i++)
                                {
                                    innerRow[column++] = null;
                                }
                                yield return innerRow;
                            }
                            else if (column < maxColNum)
                            {
                                if (_blankRow == null)
                                {
                                    _blankRow = new string[columnsCntFromFirstRowX];
                                }

                                yield return _blankRow;
                            }
                            else
                            {
                                yield return innerRow;
                            }

                            if (prevRow != -1 && prevRow < rowNum - 1)
                            {
                                for (int i = 0; i < rowNum - prevRow - 1; i++)
                                {
                                    if (_blankRow == null)
                                    {
                                        _blankRow = new string[columnsCntFromFirstRowX];
                                    }
                                    yield return _blankRow;
                                }
                            }
                            column = 0;
                            Array.Clear(innerRow, 0, innerRow.Length);
                        }
                        prevRow = rowNum;

                        colNum = reader.columnNum;
                        if (firstRowX && minColNumX == -1)
                        {
                            minColNumX = colNum;
                        }
                        if (firstRowX && maxColNum < colNum)
                        {
                            maxColNum = colNum;
                        }
                        object valueX = null;
                        object rawValue = reader.cellValue;
                        if (reader.isSharedStringVal)
                        {
                            valueX = _sharedStringArray[(int)rawValue];
                        }
                        else
                        {
                            var styleIndex = (int)reader.xfIndex;

                            if (styleIndex == 0) // general
                            {
                                valueX = GetTypedValue(rawValue);
                            }
                            else
                            {
                                int numFormatId = _stylesCellXfsArray[styleIndex].NumFmtId;

                                if (rawValue != null && _numberFormatsTypeDic.TryGetValue(numFormatId, out Type type) && type == typeof(DateTime?))
                                {
                                    valueX = DateTime.FromOADate((double)rawValue);
                                }
                                else if (rawValue != null)
                                {
                                    valueX = GetTypedValue(rawValue);
                                }
                            }
                        }

                        if (column == 0)
                        {
                            prevColNum = -1;
                        }
                        if (colNum == prevColNum + 1 || prevColNum == -1 && colNum == minColNumX)
                        {
                            innerRow[column++] = valueX;
                        }
                        else if (prevColNum == -1 && colNum != minColNumX)
                        {
                            for (int i = minColNumX; i < colNum; i++)
                            {
                                innerRow[column++] = null;
                            }
                            innerRow[column++] = valueX;
                        }
                        else if (colNum - prevColNum > 1)
                        {
                            for (int i = 0; i < colNum - prevColNum - 1; i++)
                            {
                                innerRow[column++] = null;
                            }
                            innerRow[column++] = valueX;
                        }
                        else
                        {
                            innerRow[column++] = valueX;
                        }
                        prevColNum = colNum;
                    }
                }

                if (colNum < maxColNum)
                {
                    for (int i = colNum; i < maxColNum; i++)
                    {
                        innerRow[column++] = null;
                    }
                }
                yield return innerRow;
                _blankRow = null;
                yield break;
            }
            finally
            {
                str.Dispose();
            }
        }

        public IEnumerable<object[]> GetRowsOfSheet(string sheetName)
        {
            if (mode == 0)
            {
                return GetRowsOfXlsx(sheetName);
            }
            else
            {
                return GetRowsOfXlsb(sheetName);
            }
        }

        private ZipArchiveEntry GetArchiverEntry(string sheetName)
        {
            string id = _worksheetNameToId[sheetName];
            string location = _worksheetIdToLocation[id];
            return _xlsxArchive.GetEntry($"xl/{location}");
        }

        private static (int row, int column) GetNumbersFromAdress(string startingCellAdress)
        {
            int n1 = startingCellAdress.IndexOfAny(_digits);
            string letters = startingCellAdress.Substring(0, n1);
            int rowNumFromAdress = int.Parse(startingCellAdress.Substring(n1));
            int colNumFromAdress = _letterToColumnNum[letters] + 1;
            return (rowNumFromAdress - 1, colNumFromAdress - 1);
        }
        public string ReplaceSheetData(string sheetName, IDataReader reader, string startingCellAdress = "A1")
        {
            if (!_worksheetNameToId.TryGetValue(sheetName, out string id))
            {
                throw new Exception($"ReplaceSheetData - {sheetName} not found");
            }

            string location = _worksheetIdToLocation[id];
            var sheetEntryToReplace = _xlsxArchive.GetEntry($"xl/{location}");
            sheetEntryToReplace.Delete();

            (int startingRow, int startingColumn) = GetNumbersFromAdress(startingCellAdress);

            int _dateStyleNum = 0;

            for (int i = 0; i < _stylesCellXfsArray.Length; i++)
            {
                var item = _stylesCellXfsArray[i];
                int numFormatId = item.NumFmtId;
                if (_numberFormatsTypeDic.TryGetValue(numFormatId, out Type type) && type == typeof(DateTime?))
                {
                    _dateStyleNum = i;
                    break;
                }
            }

            string excelRangeString = "";
            var newData = _xlsxArchive.CreateEntry($"xl/{location}");
            using (StreamWriter sheetData = new FormattingStreamWriter(newData.Open(), Encoding.UTF8, CultureInfo.InvariantCulture.NumberFormat))
            {
                int cnt = XlsxWriter.WriteToExisting(sheetData, reader, dateStyleNum: _dateStyleNum, dateTimeStyleNum: _dateStyleNum, startingRow, startingColumn);
                string start = XlsxWriter._letters[startingColumn] + (startingRow + 1).ToString();
                string end = XlsxWriter._letters[startingColumn + reader.FieldCount - 1] + (startingRow + 1 + cnt).ToString();
                excelRangeString = start + ":" + end;
            }
            return excelRangeString;
        }

        private List<string> GetPivotTableList()
        {
            List<string> pivotTableList = new List<string>();
            var e = _xlsxArchive.GetEntry("[Content_Types].xml");
            using (var str = e.Open())
            {
                using var reader = XmlReader.Create(str);
                while (reader.Read())
                {
                    if (reader.Name == "Override")
                    {
                        string nm = reader.GetAttribute("ContentType");
                        if (nm == "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml")
                        {
                            pivotTableList.Add(reader.GetAttribute("PartName"));
                        }
                    }
                }
            }
            return pivotTableList;
        }
        private int GetCacheIdForPivotTable(string pivotTableName, List<string> pivotTableList)
        {
            int cacheId = -1;
            foreach (var pivotTableLocation in pivotTableList)
            {
                var e1 = _xlsxArchive.GetEntry(pivotTableLocation[1..]);
                using var str = e1.Open();
                using var reader = XmlReader.Create(str);
                while (reader.Read())
                {
                    if (reader.IsStartElement("pivotTableDefinition"))
                    {
                        string name = reader.GetAttribute("name");
                        if (name == pivotTableName)
                        {
                            int.TryParse(reader.GetAttribute("cacheId"), out cacheId);
                            break;
                        }
                    }
                }
            }

            return cacheId;
        }
        public void ReplacePivotTableDim(string pivotTableName, string referention, bool doRefreshOnLoad = true)
        {
            var pivotTableList = GetPivotTableList();
            int cacheId = GetCacheIdForPivotTable(pivotTableName, pivotTableList);

            if (cacheId == -1)
            {
                throw new Exception($"{pivotTableName} - not found");
            }

            string rId = _pivotCacheIdtoRid[cacheId];
            string tableCacheLocation = _pivotCachRidToLocation[rId];
            var pivotCacheEntry = _xlsxArchive.GetEntry("xl/" + tableCacheLocation);

            bool addRefreshOnLoad = false;
            bool replaceRefresh = false;
            string reff = null;
            string pivotTableXmlAsPlainTxt = null;
            using (var str = pivotCacheEntry.Open())
            {
                using var reader = XmlReader.Create(str);
                while (reader.Read())
                {
                    if (reader.IsStartElement("pivotCacheDefinition"))
                    {
                        string refreshOnLoad = reader.GetAttribute("refreshOnLoad");
                        if (refreshOnLoad == null)
                        {
                            addRefreshOnLoad = true;
                        }
                        else if (refreshOnLoad == "0")
                        {
                            replaceRefresh = true;
                        }
                    }
                    else if (reader.IsStartElement("cacheSource"))
                    {
                        do
                        {
                            reader.Read();
                        } while (!reader.IsStartElement("worksheetSource"));
                        reff = reader.GetAttribute("ref");
                        break;
                    }
                }
            }

            if (reff != null)
            {
                using var str = pivotCacheEntry.Open();
                using var reader = new StreamReader(str);
                pivotTableXmlAsPlainTxt = reader.ReadToEnd();
            }
            pivotCacheEntry.Delete();

            int firsPartIndex = pivotTableXmlAsPlainTxt.IndexOf(@"</cacheSource>");
            string firsPartTxt = pivotTableXmlAsPlainTxt[0..firsPartIndex];

            var ent = _xlsxArchive.CreateEntry("xl/" + tableCacheLocation);
            using (var str = ent.Open())
            {
                using var sw = new StreamWriter(str);
                if (doRefreshOnLoad && addRefreshOnLoad)
                {
                    firsPartTxt = firsPartTxt.Replace(" r:id=", " refreshOnLoad=\"1\" r:id=");
                }
                else if (doRefreshOnLoad && replaceRefresh)
                {
                    firsPartTxt = firsPartTxt.Replace("refreshOnLoad=\"0\" r:id=", "refreshOnLoad=\"1\" r:id=");
                }

                firsPartTxt = firsPartTxt.Replace($"ref=\"{reff}\"", $"ref=\"{referention}\"");

                sw.Write(firsPartTxt);
                sw.Write(pivotTableXmlAsPlainTxt.AsSpan().Slice(firsPartIndex));
            }
        }

        private string _actualSheetName;
        public override string ActualSheetName
        {
            get => _actualSheetName;
            set
            {
                _rowCount = -2;
                _actualSheetName = value;
            }
        }

        private IEnumerator<object[]> _enumerator;

        private int _resultCount = -1;
        public override int ResultsCount { get => _resultCount; }
        private string Name { get => ActualSheetName; }

        private int _rowCount = -2;

        public override int RowCount { get => _rowCount != -2 ? _rowCount : PrepareRowCnt(); }

        private int PrepareRowCnt()
        {
            _rowCount = -1;
            if (_actualSheetDimensions == null)
            {
                _rowCount = 123123123;
                return _rowCount;
            }

            int i1 = _actualSheetDimensions.IndexOf(":");
            string t1 = _actualSheetDimensions[..i1];
            int i2 = t1.IndexOfAny(_digits);
            int.TryParse(t1[i2..], out int start);

            t1 = _actualSheetDimensions[(i1 + 1)..];
            i2 = t1.IndexOfAny(_digits);

            int.TryParse(t1[i2..], out int end);
            _rowCount = end - start; // header is not row !!

            return _rowCount;
        }

        public override bool Read()
        {
            if (_enumerator == null)
            {
                _enumerator = GetRowsOfSheet(ActualSheetName).GetEnumerator();
            }
            if (!_enumerator.MoveNext())
            {
                _enumerator = null;
                return false;
            }
            return true;
        }

        public override string GetName(int i)
        {
            return GetValue(i).ToString();
        }
        public override object GetValue(int i)
        {
            return _enumerator.Current[i];
        }
        public override Type GetFieldType(int i)
        {
            return _enumerator.Current[i]?.GetType();
        }
        public override void GetValues(object[] row)
        {
            var arr = _enumerator.Current;
            for (int i = 0; i < row.Length; i++)
            {
                row[i] = arr[i];
            }
        }

        private static object GetTypedValue(object rawValue)
        {
            if (rawValue is string || rawValue is bool)
            {
                return rawValue;
            }
            else
            {
                long l1 = Convert.ToInt64(rawValue);
                double res = l1 - (double)rawValue;
                if (res < 3 * double.Epsilon && res > -3 * double.Epsilon)
                {
                    return Convert.ToInt64(rawValue);
                }
            }
            return rawValue;
        }
    }
}
