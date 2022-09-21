using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

//some code from https://github.com/MarkPflug/Sylvan.Data.Excel
//some code from https://github.com/ExcelDataReader/ExcelDataReader

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
            @"yyyy\-mm\-dd\ hh:mm:ss",
            @"yyyy\-mm\-dd;@",
            @"[$-409]dd\-mm\-yy\ h:mm\ AM/PM;@",
            @"dd\-mm\-yy\ h:mm;@",
            @"[$-415]mmmmm;@",
            @"[$-415]mmmmm\.yy;@",
            @"\-m\-yyyy;@",
            @"[$-415]d\-mmm\-yyyy;@",
            @"d\-m\-yyyy;@"
        };

        private string _sharedStringsLocation = null;
        private string _stylesLocation = null;
        private string _themeLocation = null;
        private int _uniqueStringCount = -1;
        private int _stringCount = -1;

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
        /// <summary>
        /// Initialize file to read (detection xlsx/xlsb by extension)
        /// </summary>
        /// <param name="path"></param>
        /// <param name="readSharedStrings"></param>
        /// <param name="updateMode"></param>
        public override void Open(string path, bool readSharedStrings = true, bool updateMode = false, Encoding encoding = null)
        {
            if (path.EndsWith("xlsb", StringComparison.OrdinalIgnoreCase))
            {
                mode = Modes.xlsb;
                OpenXlsb(path, readSharedStrings);
            }
            else
            {
                mode = Modes.xlsx;
                OpenXlsx(path,updateMode: updateMode); 
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

        /// <summary>
        /// xlsb read strategy, true = More RAM needed but faster
        /// </summary>
        public bool UseMemoryStreamInXlsb = true;
        private static MemoryStream GetMemoryStream(Stream streamToRead, long length)
        {
            byte[] byteArray = new byte[length];
            int pos = 0;
            int bytesRead = 0;
            int toRead = 65_536;
            while (true)
            {
                if (length - bytesRead < toRead)
                {
                    toRead = (int)length - bytesRead;
                }

                pos = streamToRead.Read(byteArray, bytesRead, toRead);
                bytesRead += pos;
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

            var sharedstringsEntry = _xlsxArchive.GetEntry($@"xl/{_sharedStringsLocation}");
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

                if (_uniqueStringCount !=-1)
                {
                    _sharedStringArray = new string[_uniqueStringCount];
                }
                else
                {
                    _sharedStringArray = new string[1024];
                }

                int stringNum = 0;
                while (reader.Read())
                {
                    if (reader.IsStartElement("si"))
                    {
                        reader.Read();
                        if (reader.IsStartElement("t"))
                        {
                            string vall = reader.ReadElementContentAsString();
                            // si -> t->  t -> value
                            //string vall = reader.Value;

                            if (stringNum >= _sharedStringArray.Length)
                            {
                                Array.Resize(ref _sharedStringArray, 2 * _sharedStringArray.Length);
                            }

                            _sharedStringArray[stringNum++] = vall;
                        }
                    }
                }

                if (stringNum != _sharedStringArray.Length)
                {
                    Array.Resize(ref _sharedStringArray, stringNum);
                }

            }
            finally
            {
                str.Dispose();
            }

        }

        public void FillBinSharedStrings()
        {
            if (_sharedStringsLocation == null)
            {
                throw new Exception("no shared strings found");
            }

            var sharedstringsEntry = _xlsxArchive.GetEntry($@"xl/{_sharedStringsLocation}");

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
                int stringNum = 0;

                reader.ReadSharedStrings();
                if (_sharedStringArray == null && reader.sharedStringUniqueCount != 0)
                {
                    _sharedStringArray = new string[reader.sharedStringUniqueCount];
                }
                else
                {
                    _sharedStringArray = new string[1024];
                }

                while (reader.ReadSharedStrings())
                {
                    string vall = reader.sharedStringValue;
                    if (vall != null)
                    {
                        //_sharedStringList.Add(vall);
                        if (stringNum >= _sharedStringArray.Length)
                        {
                            Array.Resize(ref _sharedStringArray, 2 * _sharedStringArray.Length);
                        }
                        _sharedStringArray[stringNum++] = vall;
                    }
                }
                if (stringNum != _sharedStringArray.Length)
                {
                    Array.Resize(ref _sharedStringArray, stringNum);
                }

            }
            finally
            {
                str.Dispose();
            }

        }

        private void FillStyles()
        {
            var _stylesCellXfs = new List<StyleInfo>();

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
                                _stylesCellXfs.Add(new StyleInfo() { XfId = xfId, NumFmtId = numFmtId/*, ApplyNumberFormat = applyNumberFormat*/ });
                                reader.Skip();
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
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
        }

        private void FillBinStyles()
        {
            var _stylesCellXfs = new List<StyleInfo>();

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
        }

        public override string[] GetScheetNames()
        {
            return _worksheetIdToName.Values.ToArray();
        }

        private static readonly char[] _digits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

        private string _actualSheetDimensions = null;

        Stream sheetStream;
        XmlReader xmlReader;
        BiffReaderWriter biffReader;
        ZipArchiveEntry sheetEntry;

        int prevRowNum = -1;
        bool isFirstRow = true;
        int columnsCntFromFirstRow = -1;
        int minColNum = -1;
        int maxColNum = -1;
        int collDif = 0;
        int colNum = -1;
        int prevColNum = -1;
        int howManyEmptyRow = -1;
        int emptyRowCnt = -1;
        bool success = false;
        int len = 0;
        int rowNum = 0;
        bool returnValue = true;
        void sheetPreInitialize()
        {
            prevRowNum = -1;
            //isFirstRow = true;
            columnsCntFromFirstRow = -1;
            minColNum = -1;
            maxColNum = -1;
            collDif = 0;
            colNum = -1;
            prevColNum = -1;
            howManyEmptyRow = -1;
            emptyRowCnt = -1;
            success = false;
            len = 0;
            rowNum = 0;
            returnValue = true;
            innerRow = new FieldInfo[4096];
            sheetEntry = GetArchiverEntry(ActualSheetName);
        }
        private void initSheetXlsxReader()
        {
            sheetPreInitialize();
            //sheetStream = sheetEntry.Open();
            sheetStream = new BufferedStream(sheetEntry.Open(), 65_536);
            xmlReader = XmlReader.Create(sheetStream, _xmlSettings);
            xmlReader.Read();
            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement("sheetData"))
                {
                    prevRowNum = -1;
                    return;
                }
                else if (xmlReader.IsStartElement("dimension"))
                {
                    _actualSheetDimensions = xmlReader.GetAttribute("ref");
                    xmlReader.Skip();
                }
                else if (xmlReader.Depth == 0)
                {
                    xmlReader.Read();
                }
                else
                {
                    xmlReader.Skip();
                }
            }
        }

        private bool ReadXlsx()
        {
            if (isFirstRow)
            {
                initSheetXlsxReader();
            }
            if (emptyRowCnt < howManyEmptyRow)
            {
                emptyRowCnt++;
                prevRowNum++;
                return true;
            }

            if (emptyRowCnt == -1)
            {
                if (!xmlReader!.ReadToFollowing("row"))
                {
                    xmlReader.Dispose();
                    sheetStream.Dispose();
                    isFirstRow = true; // RESET
                    return false;
                }

                success = xmlReader.MoveToAttribute("r");
                len = xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                rowNum = ParseToUnsignedIntFromBuffer(_buffer, len);
            }
            howManyEmptyRow = -1;
            emptyRowCnt = -1;

            //empty row/s
            if (prevRowNum != -1 && rowNum > prevRowNum + 1)
            {
                for (int i = 0; i < columnsCntFromFirstRow; i++)
                {
                    innerRow[i].type = ExcelDataType.Null;
                }
                howManyEmptyRow = rowNum - prevRowNum - 1;
                prevRowNum++;
                emptyRowCnt = 1;
                return true;
            }

            prevRowNum = rowNum;
            while (xmlReader.Read() && xmlReader.IsStartElement("c"))
            {
                //reader.MoveToAttribute("r");
                xmlReader.MoveToNextAttribute();

                len = xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                colNum = -1;

                //Sylvan
                for (int j = 0; j < len; j++)
                {
                    char c = _buffer[j];
                    if (c < 'A' || c > 'Z')
                    {
                        break;
                    }
                    int v = c - 'A';
                    if ((uint)v < 26u)
                    {
                        colNum = ((colNum + 1) * 26) + v;
                    }
                }
                colNum++;

                if (isFirstRow && minColNum == -1)
                {
                    minColNum = colNum;
                    collDif = minColNum - 1;
                }
                if (isFirstRow && maxColNum < colNum)
                {
                    maxColNum = colNum;
                }

                ref FieldInfo valueX = ref innerRow[colNum - 1 - collDif];

                bool isEmptyElement = xmlReader.IsEmptyElement;
                if (!isEmptyElement)
                {
                    char sstMark = '\0';
                    int sstLen = 0;
                    int styleId = -1;

                    success = xmlReader.MoveToNextAttribute();

                    if (success && xmlReader.Name == "s")
                    {
                        len = xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                        styleId = ParseToUnsignedIntFromBuffer(_buffer, len);
                        success = xmlReader.MoveToNextAttribute();
                    }

                    if (success && xmlReader.Name == "t")
                    {
                        sstLen = xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                        sstMark = _buffer[0];
                    }

                    xmlReader.MoveToElement();
                    //bool success = false;
                    if (sstMark == 'i' && sstLen == 9)
                    {
                        //success =
                        xmlReader.ReadToDescendant("is");
                        xmlReader.Read();
                        xmlReader.Read();
                        valueX.type = ExcelDataType.String;
                        valueX.strValue = xmlReader.ReadContentAsString();
                    }
                    else if (xmlReader.ReadToDescendant("v"))
                    {
                        xmlReader.Read();

                        if (sstMark == 's' && sstLen == 1) // 's' = string/sharedstring'b' = boolean, 'e' = error
                        {
                            len = xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                            valueX.type = ExcelDataType.String;
                            valueX.strValue = _sharedStringArray[ParseToUnsignedIntFromBuffer(_buffer, len)];
                        }
                        else if (sstMark == 's' || sstMark == 'i' && sstLen == 9)  // InlineStr?
                        {
                            if (xmlReader.NodeType != XmlNodeType.EndElement || xmlReader.Name != "c")
                            {
                                //if (xmlReader.NodeType != XmlNodeType.Text)
                                //{ 
                                //    xmlReader.Read();
                                //}
                                valueX.type = ExcelDataType.String;
                                valueX.strValue = xmlReader.ReadContentAsString();
                            }
                            else
                            {
                                valueX.type = ExcelDataType.Null;
                            }
                            //valueX = _buffer.AsSpan(0, len).ToString();
                        }
                        else
                        {
                            len = xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                            if (sstMark == 'b') // 'b' = boolean, 'e' = error
                            {
                                valueX.type = ExcelDataType.Boolean;
                                valueX.int64Value = (_buffer[0] - '0');
                            }
                            else if (sstMark == 'e') // 'b' = boolean, 'e' = error
                            {
                                valueX.type = ExcelDataType.String;
                                valueX.strValue = "error in cell";
                            }
                            else if (styleId != -1)
                            {
                                var s = _stylesCellXfsArray[styleId];
                                int numFormatId = s.NumFmtId;

                                if (_numberFormatsTypeDic.TryGetValue(numFormatId, out Type type) && type == typeof(DateTime?)
                                    && double.TryParse(_buffer.AsSpan(0, len), NumberStyles.Any, provider: invariantCultureInfo, out double doubleDate)
                                    )
                                {
                                    valueX.type = ExcelDataType.DateTime;
                                    valueX.dtValue = DateTime.FromOADate(doubleDate);
                                }
                                else
                                {
                                    if (ContainsDoubleMarks(_buffer, len))
                                    {
                                        valueX.type = ExcelDataType.Double;
                                        valueX.doubleValue = double.Parse(_buffer.AsSpan(0, len), provider: invariantCultureInfo);
                                    }
                                    else
                                    {
                                        valueX.type = ExcelDataType.Int64;
                                        valueX.int64Value = ParseToInt64FromBuffer(_buffer, len);
                                    }
                                }
                            }
                            else
                            {
                                if (ContainsDoubleMarks(_buffer, len))
                                {
                                    valueX.type = ExcelDataType.Double;
                                    valueX.doubleValue = double.Parse(_buffer.AsSpan(0, len), provider: invariantCultureInfo);
                                }
                                else
                                {
                                    valueX.type = ExcelDataType.Int64;
                                    valueX.int64Value = ParseToInt64FromBuffer(_buffer, len);
                                }
                            }
                        }
                    }
                    else
                    {
                        valueX.type = ExcelDataType.Null;
                    }
                }

                if (prevColNum >= colNum && colNum > minColNum)
                {
                    for (int i = minColNum; i < colNum; i++)
                    {
                        innerRow[i - 1 - collDif].type = ExcelDataType.Null;
                    }
                }
                else if (colNum - prevColNum > 1 && prevColNum != -1)
                {
                    for (int i = 0; i < colNum - prevColNum - collDif - 1; i++)
                    {
                        innerRow[prevColNum - collDif + i].type = ExcelDataType.Null;
                    }
                }
                prevColNum = colNum;

                if (!isEmptyElement) // depth = ...
                {
                    while (xmlReader.Depth > 3)
                    {
                        //reader.Skip();
                        xmlReader.Read();
                    }
                }
            }

            if (isFirstRow)
            {
                isFirstRow = false;
                columnsCntFromFirstRow = maxColNum - minColNum + 1;
                Array.Resize<FieldInfo>(ref innerRow, columnsCntFromFirstRow);
                FieldCount = columnsCntFromFirstRow;
            }
            if (colNum < maxColNum)
            {
                for (int i = colNum; i < maxColNum; i++)
                {
                    innerRow[i - collDif].type = ExcelDataType.Null;
                }
            }

            return true;
        }



        private void initSheetXlsbReader()
        {
            sheetPreInitialize();

            if (UseMemoryStreamInXlsb)
            {
                sheetStream = GetMemoryStream(sheetEntry.Open(), sheetEntry.Length);
            }
            else
            {
                sheetStream = new BufferedStream(sheetEntry.Open());
            }
            biffReader = new BiffReaderWriter(sheetStream);

            while (!biffReader.readCell) //read to cell
            {
                biffReader.ReadWorksheet();
            }
            //prevColNum = colNum;
            rowNum = biffReader.rowIndex;
            colNum = biffReader.columnNum;
        }

        private bool ReadXlsb()
        {
            //fist time = initialize
            if (isFirstRow)
            {
                initSheetXlsbReader();
            }

            //last row is not complete
            if (!isFirstRow && colNum > minColNum)
            {
                for (int i = minColNum; i < colNum; i++)
                {
                    innerRow[i - minColNum].type = ExcelDataType.Null;
                }
            }
            //previous read = false
            if (!returnValue)
            {
                biffReader.Dispose();
                sheetStream.Dispose();
                isFirstRow = true; // RESET
                return false;
            }

            //missing rows
            if (prevRowNum != -1 && rowNum > prevRowNum + 1)
            {
                for (int i = 0; i < columnsCntFromFirstRow; i++)
                {
                    innerRow[i].type = ExcelDataType.Null;
                }
                howManyEmptyRow = rowNum - prevRowNum - 1;
                prevRowNum++;
                emptyRowCnt = 1;
                return true;
            }
            prevRowNum = rowNum;

            while (rowNum == prevRowNum && returnValue)
            {
                if (biffReader.readCell)
                {
                    //determine first row length
                    if (isFirstRow)
                    {
                        if (minColNum == -1)
                        {
                            minColNum = colNum;
                        }
                        maxColNum = colNum;
                        columnsCntFromFirstRow = maxColNum - minColNum + 1;
                        FieldCount = columnsCntFromFirstRow;
                    }

                    ref FieldInfo valueX = ref innerRow[colNum - minColNum];

                    valueX.type = ExcelDataType.String;
                    switch (biffReader.cellType)
                    {
                        case CellType.sharedString:
                            valueX.type = ExcelDataType.String;
                            valueX.strValue = _sharedStringArray[biffReader.intValue];
                            break;
                        case CellType.stringVal:
                            valueX.type = ExcelDataType.String;
                            valueX.strValue = biffReader.stringValue;
                            break;
                        case CellType.boolVal:
                            valueX.type = ExcelDataType.Boolean;
                            valueX.int64Value = biffReader.boolValue ? 1 : 0;
                            break;
                        case CellType.doubleVal:
                            {
                                double doubleVal = biffReader.doubleVal;
                                var styleIndex = biffReader.xfIndex;
                                if (styleIndex == 0) // general
                                {
                                    SetValueForXlsb(doubleVal, ref valueX);
                                }
                                else
                                {
                                    int numFormatId = _stylesCellXfsArray[styleIndex].NumFmtId;
                                    if (_numberFormatsTypeDic.TryGetValue(numFormatId, out Type type) && type == typeof(DateTime?))
                                    {
                                        valueX.type = ExcelDataType.DateTime;
                                        valueX.dtValue = DateTime.FromOADate((double)doubleVal);
                                    }
                                    else
                                    {
                                        SetValueForXlsb(doubleVal, ref valueX);
                                    }
                                }
                                break;
                            }
                        default:
                            valueX.type = ExcelDataType.Null;
                            break;
                    }
                }

                biffReader.ReadWorksheet();
                while (returnValue && !biffReader.readCell)
                {
                    returnValue = biffReader.ReadWorksheet();
                }

                prevColNum = colNum;
                rowNum = biffReader.rowIndex;
                colNum = biffReader.columnNum;

                if (!isFirstRow)
                {
                    if (colNum > prevColNum + 1 && rowNum == prevRowNum)
                    {
                        for (int i = prevColNum + 1; i < colNum; i++)
                        {
                            innerRow[i - minColNum].type = ExcelDataType.Null;
                        }
                    }
                    else if (rowNum > prevRowNum && prevColNum < maxColNum)
                    {
                        for (int i = 1; i <= maxColNum - prevColNum; i++)
                        {
                            innerRow[prevColNum + i - minColNum].type = ExcelDataType.Null;
                        }
                    }
                }
            }

            if (isFirstRow)
            {
                isFirstRow = false;
                Array.Resize(ref innerRow, columnsCntFromFirstRow);
            }

            return true;
        }

        public override bool Read()
        {
            if (mode == Modes.xlsx)
            {
                return ReadXlsx();
            }
            else
            {
                return ReadXlsb();
            }
        }

        //inspired by https://github.com/MarkPflug/Sylvan.Data.Excel
        private readonly static char[] _buffer = new char[64];

        private static int ParseToUnsignedIntFromBuffer(char[] buff, int len)
        {
            int res = 0;
            for (int i = 0; i < len; i++)
            {
                res = res * 10 + (buff[i] - '0');
            }
            return res;
        }

        private static Int64 ParseToInt64FromBuffer(char[] buff, int len)
        {
            Int64 res = 0;
            int start = buff[0] == '-' ? 1 : 0;
            for (int i = start; i < len; i++)
            {
                res = res * 10 + (buff[i] - '0');
            }
            return start == 1 ? -res : res;
        }

        private static bool ContainsDoubleMarks(char[] buff, int len)
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

        public IEnumerable<object[]> GetRowsOfXlsx(string sheetName)
        {
            ActualSheetName = sheetName;
            Read();
            object[] row = new object[FieldCount];
            GetValues(row);
            yield return row;

            while (ReadXlsx())
            {
                GetValues(row);
                yield return row;
            }
        }

        public IEnumerable<object[]> GetRowsOfXlsb(string sheetName)
        {
            ActualSheetName = sheetName;
            Read();
            object[] row = new object[FieldCount];
            GetValues(row);
            yield return row;

            while (ReadXlsb())
            {
                GetValues(row);
                yield return row;
            }
        }

        public IEnumerable<object[]> GetRowsOfSheet(string sheetName)
        {
            if (mode == Modes.xlsx)
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

        /// <summary>
        /// clear and fill existing tab with dataReader
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="reader"></param>
        /// <param name="startingCellAdress"></param>
        /// <returns></returns>
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

        /// <summary>
        /// changes data reference for pivot table, use with ReplaceSheetData
        /// </summary>
        /// <param name="pivotTableName">name of pivot table</param>
        /// <param name="referention">new reference</param>
        /// <param name="doRefreshOnLoad">refresh pivot table on open</param>
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

        private int _resultCount = -1;
        public override int ResultsCount { get => _resultCount; }
        private string Name { get => ActualSheetName; }

        private int _rowCount = -2;

        /// <summary>
        /// row count (not always available)
        /// </summary>
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

        private void SetValueForXlsb(double rawValue, ref FieldInfo fieldInfo)
        {
            long l1 = Convert.ToInt64(rawValue);
            double res = l1 - /*(double)*/rawValue;
            if (res < 3 * double.Epsilon && res > -3 * double.Epsilon)
            {
                fieldInfo.type = ExcelDataType.Int64;
                fieldInfo.int64Value = l1;
            }
            else
            {
                fieldInfo.type = ExcelDataType.Double;
                fieldInfo.doubleValue = rawValue;
            }
        }
    }
}
