using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace SpreadSheetTasks
{
    public class XlsxWriter : ExcelWriter, IDisposable
    {
        private readonly bool _inMemoryMode;
        private readonly int _bufferSize;
        internal static readonly string[] _letters;
        private readonly bool _useScharedStrings;

        private readonly CompressionLevel clvl;

        private const double _dateTimeWidth = 16.0;
        private const double _dateWidth = 10.140625;


        public XlsxWriter(string filePath, int bufferSize = 4096, bool InMemoryMode = true, bool useScharedStrings = true, CompressionLevel _clvl = CompressionLevel.Optimal)
            : this(new FileStream(filePath, FileMode.Create), bufferSize, InMemoryMode, useScharedStrings, _clvl, leaveExcelArchiveOpen:false)
        {
            _excelStreamWasProvided = false;
        }

        public XlsxWriter(Stream stream, int bufferSize = 4096, bool InMemoryMode = true, bool useScharedStrings = true, CompressionLevel _clvl = CompressionLevel.Optimal, bool leaveExcelArchiveOpen = true) 
        {
            _excelStreamWasProvided = true;
            _newExcelFileStream = stream;
            _bufferSize = bufferSize;
            _inMemoryMode = InMemoryMode;
            TryToSpecifyWidthForMemoryMode = InMemoryMode;

            _useScharedStrings = useScharedStrings;
            if (_useScharedStrings)
            {
                _sstDic = new Dictionary<string, int>();
            }
            clvl = _clvl;

            try
            {
                _newExcelFileStream = stream;
                _excelArchiveFile = new ZipArchive(_newExcelFileStream, ZipArchiveMode.Create,leaveOpen:true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }


        static XlsxWriter()
        {
            var lettersX = new List<string>();

            for (int i = 65; i < 91; i++)
            {
                lettersX.Add(((char)i).ToString());
            }

            var tempList = new List<string>();
            foreach (var item in lettersX)
            {

                for (int i = 65; i < 91; i++)
                {
                    tempList.Add(item + ((char)i).ToString());
                }
            }

            var tempList2 = new List<string>();
            foreach (var item in tempList)
            {
                if (item.CompareTo("XY") > 0)
                {
                    break;
                }

                for (int i = 65; i < 91; i++)
                {
                    tempList2.Add(item + ((char)i).ToString());
                }
            }

            lettersX.AddRange(tempList);
            lettersX.AddRange(tempList2);
            _letters = lettersX.ToArray();
        }
        private string GetTempFileFullPath()
        {
            return $"{Path.GetTempPath()}\\{Path.GetRandomFileName()}";
        }
        public override void Save()
        {
            DoOnCompress();
            if (!_inMemoryMode)
            {
                for (int i = 0; i < sheetCnt + 1; i++)
                {
                    _excelArchiveFile.CreateEntryFromFile(_sheetList[i].pathOnDisc, _sheetList[i].pathInArchive, clvl);
                    File.Delete(_sheetList[i].pathOnDisc);
                }
            }
            base.Save();
        }
        public override void AddSheet(string sheetName, bool hidden = false)
        {
            string archveSheetName = "sheet" + (sheetCnt + 2);
            _sheetList.Add((sheetName, String.Format(@"xl/worksheets/{0}.xml", archveSheetName), GetTempFileFullPath(), hidden, archveSheetName, (sheetCnt + 2),null));
            //_sheetList.Add((sheetName, String.Format(@"xl/worksheets/{0}.xml", sheetName), getTempFileFullPath(), hidden, sheetName));
        }
        public override void WriteSheet(IDataReader dataReader, Boolean headers = true, int overLimit = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            sheetCnt++;
            this.areHeaders = headers;
            _dataColReader = new DataColReader(dataReader, headers, overLimit);
            if (_inMemoryMode)
            {
                var e1 = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt].pathInArchive, clvl);
                using StreamWriter daneZakladkiX = new FormattingStreamWriter(e1.Open(), Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _rowsCount = WriteSheet(daneZakladkiX, startingRow, startingColumn, doAutofilter: doAutofilter) - 1;
            }
            else
            {
                using StreamWriter sheedData = new FormattingStreamWriter(_sheetList[sheetCnt].pathOnDisc, false, Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _rowsCount = WriteSheet(sheedData, startingRow, startingColumn, doAutofilter: doAutofilter) - 1;
            }
        }
        public override void WriteSheet(DataTable dataTable, Boolean headers = true, int overLimit = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            sheetCnt++;
            this.areHeaders = headers;
            _dataColReader = new DataColReader(dataTable, headers, overLimit);
            if (_inMemoryMode)
            {
                var e1 = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt].pathInArchive, clvl);
                using StreamWriter daneZakladkiX = new FormattingStreamWriter(e1.Open(), Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _rowsCount = WriteSheet(daneZakladkiX, startingRow, startingColumn,doAutofilter: doAutofilter) - 1;
            }
            else
            {
                using StreamWriter daneZakladkiX = new FormattingStreamWriter(_sheetList[sheetCnt].pathOnDisc, false, Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _rowsCount = WriteSheet(daneZakladkiX, startingRow, startingColumn, doAutofilter: doAutofilter) - 1;
            }
        }

        private static readonly CultureInfo _invariantCulture = CultureInfo.InvariantCulture;
        private const int BUFFER_SIZE = 65_536;
        private const int BUFFER_SIZE_HALF = BUFFER_SIZE / 2;
        private readonly char[] buffer = new char[BUFFER_SIZE];
        private int currentBufferOffset = 0;

        private void writeByteToBuffer(byte val)
        {
            val.TryFormat(buffer.AsSpan(currentBufferOffset), out int len, default, _invariantCulture);
            currentBufferOffset += len;
        }
        private void writesByteToBuffer(sbyte val)
        {
            val.TryFormat(buffer.AsSpan(currentBufferOffset), out int len, default, _invariantCulture);
            currentBufferOffset += len;
        }

        private void writeInt16ToBuffer(Int16 val)
        {
            val.TryFormat(buffer.AsSpan(currentBufferOffset), out int len, default, _invariantCulture);
            currentBufferOffset += len;
        }

        private void writeInt32ToBuffer(Int32 val)
        {
            val.TryFormat(buffer.AsSpan(currentBufferOffset), out int len, default, _invariantCulture);
            currentBufferOffset += len;
        }
        private void writeInt64ToBuffer(Int64 val)
        {
            val.TryFormat(buffer.AsSpan(currentBufferOffset), out int len, default, _invariantCulture);
            currentBufferOffset += len;
        }
        private void writeDoubleToBuffer(double val)
        {
            val.TryFormat(buffer.AsSpan(currentBufferOffset), out int len, default, _invariantCulture);
            currentBufferOffset += len;
        }
        private void writeFloatToBuffer(float val)
        {
            val.TryFormat(buffer.AsSpan(currentBufferOffset), out int len, default, _invariantCulture);
            currentBufferOffset += len;
        }

        //private void writeStringToBuffer(string val)
        //{
        //    for (int i = 0; i < val.Length; i++)
        //    {
        //        char c = val[i];
        //        buffer[currentBufferOffset++] = c;
        //    }
        //}
        //[MethodImpl(MethodImplOptions.AggressiveOptimization)]
        private void writeStringToBuffer(ReadOnlySpan<char> val)
        {
            val.CopyTo(buffer.AsSpan(currentBufferOffset));
            currentBufferOffset += val.Length;
        }

        public bool TryToSpecifyWidthForMemoryMode { get; set; }
        private int WriteSheet(StreamWriter sheetWritter, int startingRow, int startingColumn, bool doAutofilter = false)
        {
            if (doAutofilter)
            {
                autofilterIsOn = true;
            }
            int rowNum = 0;

            int ColumnCount = _dataColReader.FieldCount;
            colWidesArray = new double[ColumnCount];
            Array.Fill<double>(colWidesArray, -1.0);

            typesArray = new int[ColumnCount];
            newTypes = new TypeCode[ColumnCount];

            if (_inMemoryMode)
            {
                sheetWritter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sheetWritter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

                if (TryToSpecifyWidthForMemoryMode && _dataColReader.dataReader != null)
                {
                    var rdr = _dataColReader.dataReader;
                    for (int l = 1; l <= ColumnCount; l++)
                    {
                        int lenn = rdr.GetName(l - 1).Length + (doAutofilter ?2:0);
                        double tempWidth = 1.25 * lenn + 2;
                        if (tempWidth > _MAX_WIDTH)
                        {
                            tempWidth = _MAX_WIDTH;
                        }
                        if (colWidesArray[l - 1] < tempWidth)
                        {
                            colWidesArray[l - 1] = tempWidth;
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

                        for (int i = 0; i < arr.Length; i++)
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
                        SetColsLengtth(ColumnCount, arr);
                    }
                    areNextRows = rdr.Read();
                    _dataColReader.AreNextRows = areNextRows;

                    sheetWritter.Write("<cols>");
                    for (int l = 1; l <= ColumnCount; l++)
                    {
                        sheetWritter.Write(String.Format(CultureInfo.InvariantCulture.NumberFormat, "<col min=\"{0}\" max=\"{0}\" width=\"{1}\" bestFit = \"1\" customWidth=\"1\" />", l + startingColumn, colWidesArray[l - 1]));
                    }
                    sheetWritter.Write("</cols>");
                }
                else if (TryToSpecifyWidthForMemoryMode && _dataColReader._dataTable != null)
                {
                    sheetWritter.Write($"<dimension ref=\"{_letters[startingColumn]}{startingRow + 1}:{_letters[ColumnCount - 1 + startingColumn]}{_dataColReader.DataTableRowsCount + 1 + startingRow}\"/>");

                    _dataColReader.GetWidthFromDataTable(colWidesArray, _MAX_WIDTH, doAutofilter);
                    sheetWritter.Write("<cols>");
                    for (int l = 1; l <= ColumnCount; l++)
                    {
                        sheetWritter.Write(String.Format(CultureInfo.InvariantCulture.NumberFormat, "<col min=\"{0}\" max=\"{0}\" width=\"{1}\" bestFit = \"1\" customWidth=\"1\" />", l + startingColumn, colWidesArray[l - 1]));
                    }
                    sheetWritter.Write("</cols>");
                }
                sheetWritter.Write("<sheetData>");
                colWidesArray = null;
            }
            else
            {
                for (int i = 0; i < 600 + ColumnCount * 100; i++)
                {
                    sheetWritter.Write(" ");
                }
                sheetWritter.WriteLine();
            }

            while (_dataColReader.Read())
            {
                if (rowNum == 0 || this.areHeaders && rowNum == 1)
                {
                    if (rowNum == 0 && this.areHeaders)
                    {
                        for (int i = 0; i < ColumnCount; i++)
                        {
                            typesArray[i] = 0;
                            newTypes[i] = TypeCode.String;
                        }
                    }
                    else
                    {
                        ExcelWriter.SetTypes(_dataColReader, typesArray, newTypes, ColumnCount);
                    }
                }

                writeStringToBuffer("<row r=\"");
                writeInt32ToBuffer(rowNum + 1 + startingRow);
                writeStringToBuffer("\">");

                WriteRow(rowNum, ColumnCount, startingColumn, startingRow);

                rowNum++;
                writeStringToBuffer("</row>");
                if (currentBufferOffset >= BUFFER_SIZE_HALF)
                {
                    sheetWritter.Write(buffer, 0, currentBufferOffset);
                    currentBufferOffset = 0;
                }
                if (rowNum % 10000 == 0)
                {
                    DoOn10k(rowNum);
                }
            }
            if (currentBufferOffset > 0)
            {
                sheetWritter.Write(buffer, 0, currentBufferOffset);
                currentBufferOffset = 0;
            }

            sheetWritter.Write("</sheetData>");
            if (doAutofilter)
            {
                (string name, string pathInArchive, string pathOnDisc, bool isHidden, string nameInArchive, int sheetId, string _) = this._sheetList[^1];
                this._sheetList[^1] = (name, pathInArchive, pathOnDisc, isHidden, nameInArchive, sheetId, $"{name}!${_letters[startingColumn]}${startingRow + 1}:${_letters[ColumnCount - 1 + startingColumn]}${rowNum}");

                sheetWritter.Write($"<autoFilter ref=\"{_letters[startingColumn]}{startingRow + 1}:{_letters[ColumnCount - 1 + startingColumn]}{_dataColReader.DataTableRowsCount + 1 + startingRow}\"/>");
            }
            sheetWritter.Write("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>");

            //System.NotSupportedException: 'This stream from ZipArchiveEntry does not support seeking.'
            if (!_inMemoryMode)
            {
                sheetWritter.Flush();
                sheetWritter.BaseStream.Position = 0;
                sheetWritter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sheetWritter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                sheetWritter.Write(String.Format("<dimension ref=\"A1:{0}{1}\"/>", _letters[ColumnCount - 1], rowNum + 1));

                if (doAutofilter)
                {
                    if (sheetCnt == 0)
                    {
                        sheetWritter.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\" /> <selection pane=\"bottomLeft\" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    }
                    else
                    {
                        sheetWritter.Write("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\" /> <selection pane=\"bottomLeft\" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    }
                }
                else
                {
                    if (sheetCnt == 0)
                    {
                        sheetWritter.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    }
                    else
                    {
                        sheetWritter.Write("<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    }
                }

                List<double> colWidth2 = colWidesArray.ToArray().ToList();//??
                List<double> colWidth3 = colWidth2.FindAll(x => x != 1.0);

                if (colWidth3.Count > 0)
                {
                    sheetWritter.Write("<cols>");
                    int l = 1;
                    foreach (var item in colWidth2)
                    {
                        if (item == 1.0)
                        {
                            l++;
                            continue;
                        }
                        sheetWritter.Write(String.Format(CultureInfo.InvariantCulture.NumberFormat, "<col min=\"{0}\" max=\"{0}\" width=\"{1}\" bestFit = \"1\" customWidth=\"1\" />", l + startingColumn, Math.Min(item, _MAX_WIDTH)));
                        l++;
                    }
                    sheetWritter.Write("</cols>");
                }
                sheetWritter.Write("<sheetData>");
            }

            return rowNum;


        }
        public override void WriteSheet(string[] oneColumn)
        {
            sheetCnt++;
            string[,] dane = new string[oneColumn.Length, 1];

            for (int i = 0; i < oneColumn.Length; i++)
            {
                dane[i, 0] = oneColumn[i];
            }

            _dataColReader = new DataColReader(dane);
            if (_inMemoryMode)
            {
                var e1 = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt].pathInArchive, clvl);
                using StreamWriter writter = new FormattingStreamWriter(e1.Open(), Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                WriteSheet(writter, 0, 0);
            }
            else
            {
                using StreamWriter writter = new FormattingStreamWriter(_sheetList[sheetCnt].pathOnDisc, false, Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                WriteSheet(writter, 0, 0);
            }
        }
        private void WriteRow(int rowNum, int columnCount, int startingColumn, int startingRow)
        {
            for (int column = 0; column < columnCount; column++)
            {
                if (_dataColReader.IsDBNull(column))
                    continue;

                if (newTypes[column] == TypeCode.String || newTypes[column] == TypeCode.Object || typesArray[column] == 5) // string
                {
                    string? stringValue = null;
                    if (newTypes[column] == TypeCode.String)
                    {
                        stringValue = _dataColReader.GetString(column);
                    }
                    else if (typesArray[column] == 5)
                    {
                        stringValue = Encoding.UTF8.GetString(((Memory<byte>)_dataColReader.GetValue(column)).Span);
                    }
                    else
                    {
                        stringValue = _dataColReader.GetValue(column).ToString()!;
                    }                    

                    if (!_inMemoryMode)
                    {
                        if (colWidesArray[column] < stringValue.Length * 1.25 + 2.0)
                        {
                            colWidesArray[column] = stringValue.Length * 1.25 + 2.0;
                        }
                    }

                    if (stringValue.Contains('&'))
                    {
                        stringValue = stringValue.Replace("&", "&amp;");
                    }
                    if (stringValue.Contains('<'))
                    {
                        stringValue = stringValue.Replace("<", "&lt;");
                    }
                    if (stringValue.Contains('>'))
                    {
                        stringValue = stringValue.Replace(">", "&gt;");
                    }
                    if (stringValue.Contains('\"'))
                    {
                        stringValue = stringValue.Replace("\"", "&quot;");
                    }
                    if (stringValue.Contains('\''))
                    {
                        stringValue = stringValue.Replace("'", "&apos;");
                    }

                    if (!_useScharedStrings)
                    {
                        writeStringToBuffer("<c r=\"");
                        writeStringToBuffer(_letters[column + startingColumn]);
                        writeInt32ToBuffer((rowNum + 1 + startingRow));
                        if (stringValue.Length > 0 && (stringValue[0] == ' ' || stringValue[0] == '\t'))
                        {
                            writeStringToBuffer("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                        }
                        else
                        {
                            writeStringToBuffer("\" t=\"inlineStr\"><is><t>");
                        }
                        writeStringToBuffer(stringValue);
                        writeStringToBuffer("</t></is></c>");
                    }
                    else
                    {
                        ref var dicRefValue = ref CollectionsMarshal.GetValueRefOrAddDefault(_sstDic, stringValue, out bool exists);
                        if (!exists)
                        {
                            dicRefValue = _sstCntUnique;
                            _sstCntUnique++;
                        }
                        writeStringToBuffer("<c r=\"");
                        writeStringToBuffer(_letters[column + startingColumn]);
                        writeInt32ToBuffer((rowNum + 1 + startingRow));
                        writeStringToBuffer("\" t=\"s\"><v>");
                        writeInt32ToBuffer(dicRefValue);
                        writeStringToBuffer("</v></c>");
                        _sstCntAll++;
                    }
                }
                else if (typesArray[column] == 1)//number
                {
                    writeStringToBuffer("<c r=\"");
                    writeStringToBuffer(_letters[column + startingColumn]);
                    writeInt32ToBuffer((rowNum + 1 + startingRow));
                    writeStringToBuffer("\"><v>");

                    switch (newTypes[column])
                    {
                        case TypeCode.Byte:
                            byte byteValue = _dataColReader.GetByte(column);
                            writeByteToBuffer(byteValue);
                            break;
                        case TypeCode.SByte:
                            sbyte sbyteValue = _dataColReader.GetSByte(column);
                            writesByteToBuffer(sbyteValue);
                            break;
                        case TypeCode.Int16:
                            Int16 int16Value = _dataColReader.GetInt16(column);
                            writeInt16ToBuffer(int16Value);
                            break;
                        case TypeCode.Int32:
                            Int32 int32Value = _dataColReader.GetInt32(column);
                            writeInt32ToBuffer(int32Value);
                            break;
                        case TypeCode.Int64:
                            Int64 int64Value = _dataColReader.GetInt64(column);
                            writeInt64ToBuffer(int64Value);
                            break;
                        case TypeCode.Single:
                            float doubleValue = _dataColReader.GetFloat(column);
                            writeFloatToBuffer(doubleValue);
                            break;
                        case TypeCode.Double:
                            double doubleVal = _dataColReader.GetDouble(column);
                            writeDoubleToBuffer(doubleVal);
                            break;
                        case TypeCode.Decimal:
                            decimal decimalVal = _dataColReader.GetDecimal(column);
                            writeDoubleToBuffer(decimal.ToDouble(decimalVal));
                            break;
                        default:
                            throw new Exception("number format Excel error");
                    }


                    writeStringToBuffer("</v></c>");
                }
                else if (typesArray[column] == 2) //date
                {
                    DateTime dtVal = _dataColReader.GetDateTime(column);
                    writeStringToBuffer("<c r=\"");
                    writeStringToBuffer(_letters[column + startingColumn]);
                    writeInt32ToBuffer((rowNum + 1 + startingRow));
                    writeStringToBuffer("\" s=\"1\"><v>");
                    writeDoubleToBuffer((double)(dtVal as DateTime?)?.ToOADate()!);
                    writeStringToBuffer("</v></c>");
                    if (!_inMemoryMode)
                    {
                        colWidesArray[column] = _dateWidth;
                    }
                }
                else if (typesArray[column] == 3) //datetime
                {
                    DateTime dtVal = _dataColReader.GetDateTime(column);
                    if (SuppressSomeDate && (dtVal as DateTime?).Value.Year == 1000)//1000-xx-xx
                    {
                        continue;
                    }

                    writeStringToBuffer("<c r=\"");
                    writeStringToBuffer(_letters[column + startingColumn]);
                    writeInt32ToBuffer((rowNum + 1 + startingRow));
                    writeStringToBuffer("\" s=\"2\"><v>");
                    writeDoubleToBuffer((double)((dtVal) as DateTime?)?.ToOADate()!);
                    writeStringToBuffer("</v></c>");
                    if (!_inMemoryMode)
                    {
                        colWidesArray[column] = _dateTimeWidth;
                    }
                }
                else if (newTypes[column] == TypeCode.Boolean)
                {
                    writeStringToBuffer("<c r=\"");
                    writeStringToBuffer(_letters[column + startingColumn]);
                    writeInt32ToBuffer((rowNum + 1 + startingRow));
                    writeStringToBuffer("\" t=\"b\"><v>");
                    if (_dataColReader.GetBoolean(column))
                    {
                        buffer[currentBufferOffset++] = '1';
                    }
                    else
                    {
                        buffer[currentBufferOffset++] = '0';
                    }
                    writeStringToBuffer("</v></c>");
                }
            }
        }
        private const string filterDefinedName = "_xlnm._FilterDatabase";
        internal override void FinalizeFile()
        {
            var e1 = _excelArchiveFile.CreateEntry("[Content_Types].xml", clvl);

            using (var writer = new FormattingStreamWriter(e1.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                writer.Write("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
                writer.Write("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
                if (!String.IsNullOrWhiteSpace(DocPopertyProgramName))
                {
                    writer.Write("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
                    writer.Write("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
                }
                writer.Write("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
                //if (UseDocPropsAndTheme)
                //{
                //    writer.Write("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
                //}
                writer.Write("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");

                foreach (var (_, _, _, _, nameInArchive, _, _) in _sheetList)
                {
                    writer.Write($"<Override PartName=\"/xl/worksheets/{nameInArchive}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                }
                if (_useScharedStrings)
                {
                    writer.Write("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
                }
                writer.Write("</Types>");
            }

            if (!String.IsNullOrWhiteSpace(DocPopertyProgramName))
            {
                var e2 = _excelArchiveFile.CreateEntry("docProps/app.xml", clvl);
                using var writer = new FormattingStreamWriter(e2.Open(), CultureInfo.InvariantCulture.NumberFormat);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" ");
                writer.Write("xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application");
                writer.Write($">{DocPopertyProgramName}</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs");
                writer.Write("><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt");
                writer.Write($":variant><vt:i4>{_sheetList.Count}</vt:i4></vt:variant></vt:vector></HeadingPairs>");
                writer.Write("<TitlesOfParts>");
                writer.Write($"<vt:vector size=\"{_sheetList.Count}\" baseType=\"lpstr\">");

                foreach (var (name, _, _, _, _, _,_) in _sheetList)
                {
                    writer.Write($"<vt:lpstr>{name}</vt:lpstr>");
                }
                writer.Write("</vt:vector>");
                writer.Write("</TitlesOfParts>");
                writer.Write("<Company></Company><LinksUpToDate>false</LinksUpToDate");
                writer.Write("><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000");
                writer.Write("</AppVersion></Properties>");
            }

            var e3 = _excelArchiveFile.CreateEntry("xl/workbook.xml", clvl);
            using (var writer = new FormattingStreamWriter(e3.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {

                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
                writer.Write("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><fileVersion ");
                writer.Write("appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/><workbookPr ");
                writer.Write("defaultThemeVersion=\"124226\"/><bookViews><workbookView xWindow=\"240\" yWindow=\"15\" ");
                writer.Write("windowWidth=\"16095\" windowHeight=\"9660\"/></bookViews>");
                writer.Write("<sheets>");
                for (int i = 0; i < _sheetList.Count; i++)
                {
                    var a = _sheetList[i].isHidden ? " state =\"hidden\"" : "";
                    writer.Write($"<sheet name=\"{_sheetList[i].name}\" sheetId=\"{i + 1}\"{a} r:id=\"rId{i + 1}\"/>");
                }
                writer.Write("</sheets>");

                if (autofilterIsOn)
                {
                    writer.Write("<definedNames>");
                    foreach (var item in this._sheetList)
                    {
                        if (!string.IsNullOrWhiteSpace(item.filterHeaderRange))
                        {
                            int localSheetId = item.sheetId - 1;
                            string filterHeaderRange = item.filterHeaderRange;
                            writer.Write($"<definedName name=\"{filterDefinedName}\" localSheetId=\"{localSheetId}\" hidden=\"1\">{filterHeaderRange}</definedName>");
                        }
                    }
                    writer.Write("</definedNames>");
                }

                writer.Write("<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/></workbook>");
            }

            var e4 = _excelArchiveFile.CreateEntry("xl/_rels/workbook.xml.rels", clvl);
            using (var writer = new FormattingStreamWriter(e4.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (int i = 0; i < _sheetList.Count; i++)
                {
                    writer.Write($"<Relationship Id=\"rId{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/{_sheetList[i].nameInArchive}.xml\"/>");
                }
                writer.Write($"<Relationship Id=\"rId{_sheetList.Count + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
                if (_useScharedStrings)
                {
                    writer.Write($"<Relationship Id=\"rId{_sheetList.Count + 2}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
                }
                writer.Write($"</Relationships>");
            }

            var e5 = _excelArchiveFile.CreateEntry("_rels/.rels", clvl);
            using (var writer = new FormattingStreamWriter(e5.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                writer.Write("<Relationships ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                writer.Write("<Relationship Id=\"rId1\" ");
                writer.Write("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" ");
                writer.Write("Target=\"xl/workbook.xml\"/>");

                if (!String.IsNullOrWhiteSpace(DocPopertyProgramName))
                {
                    writer.Write("<Relationship Id=\"rId2\" ");
                    writer.Write("Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" ");
                    writer.Write("Target=\"docProps/core.xml\"/>");
                    writer.Write("<Relationship Id=\"rId3\" ");
                    writer.Write("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" ");
                    writer.Write("Target=\"docProps/app.xml\"/>");
                }
                writer.Write("</Relationships>");
            }

            if (!String.IsNullOrWhiteSpace(DocPopertyProgramName))
            {
                var e6 = _excelArchiveFile.CreateEntry("docProps/core.xml", clvl);
                string stringNow = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ");

                using var writer = new FormattingStreamWriter(e6.Open(), CultureInfo.InvariantCulture.NumberFormat);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties ");
                writer.Write("xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ");
                writer.Write("xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" ");
                writer.Write("xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" ");
                writer.Write($"xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>{DocPopertyProgramName}</dc:creator>");
                writer.Write($"<cp:lastModifiedBy>{DocPopertyProgramName}</cp:lastModifiedBy>");
                writer.Write($"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{stringNow}</dcterms:created><dcterms:modified ");
                writer.Write($"xsi:type=\"dcterms:W3CDTF\">{stringNow}</dcterms:modified></cp:coreProperties>");
            }

            var e7 = _excelArchiveFile.CreateEntry("xl/styles.xml", clvl);
            using (var writer = new FormattingStreamWriter(e7.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz ");
                writer.Write("val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme ");
                writer.Write("val=\"minor\"/></font></fonts><fills count=\"2\"><fill><patternFill ");
                writer.Write("patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders ");
                writer.Write("count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs ");
                writer.Write("count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
                writer.Write("<cellXfs count=\"3\">");
                writer.Write("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
                writer.Write("<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");  // date
                writer.Write("<xf numFmtId=\"22\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");  // datetime
                writer.Write("</cellXfs>");
                writer.Write("<cellStyles ");
                writer.Write("count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>");
                writer.Write("</cellStyles>");
                writer.Write("<dxfs ");
                writer.Write("count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" ");
                writer.Write("defaultPivotStyle=\"PivotStyleLight16\"/></styleSheet>");
            }

            if (_useScharedStrings)
            {
                var entry = _excelArchiveFile.CreateEntry("xl/sharedStrings.xml", clvl);

                using var o = entry.Open();
                using var _sharedStringWritter = new FormattingStreamWriter(o, Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _sharedStringWritter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                _sharedStringWritter.Write("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
                _sharedStringWritter.Write($" count=\"{_sstCntAll}\" uniqueCount=\"{_sstDic.Count}\">");

                foreach (string dana in _sstDic.Keys)
                {
                    if (dana.Length > 0 && (dana[0] == ' ' || dana[0] == '\t'))
                    {
                        _sharedStringWritter.Write("<si><t xml:space=\"preserve\">");
                        _sharedStringWritter.Write(dana);
                        _sharedStringWritter.Write("</t></si>");
                    }
                    else
                    {
                        _sharedStringWritter.Write("<si><t>");
                        _sharedStringWritter.Write(dana);
                        _sharedStringWritter.Write("</t></si>");
                    }
                }
                _sharedStringWritter.Write("</sst>");
            }
        }
        public override void Dispose()
        {
            this.Save();
        }

        public static int WriteToExisting(StreamWriter sheetWritter, IDataReader reader, int dateStyleNum = 1, int dateTimeStyleNum = 2, int startingRow = 0, int startingColumn = 0)
        {
            sheetWritter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sheetWritter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sheetWritter.Write("<sheetData>");

            int rowNum = 0;
            int colCnt = reader.FieldCount;
            int[] typesArray = new int[colCnt];

            for (int j = 0; j < colCnt; j++)
            {
                if (ExcelWriter._stringTypes.Contains(reader.GetFieldType(j)))
                {
                    typesArray[j] = 0;
                }
                else if (ExcelWriter._numberTypes.Contains(reader.GetFieldType(j)))
                {
                    typesArray[j] = 1;
                }
                else if (reader.GetFieldType(j) == typeof(System.DateTime) && reader.GetDataTypeName(j).Equals("Date", StringComparison.OrdinalIgnoreCase))
                {
                    typesArray[j] = 2;
                }
                else if (reader.GetFieldType(j) == typeof(System.DateTime)
                    && (reader.GetDataTypeName(j).Equals("timestamp", StringComparison.OrdinalIgnoreCase) || reader.GetDataTypeName(j).Equals("DateTime", StringComparison.OrdinalIgnoreCase)))
                {
                    typesArray[j] = 3;
                }
                else // Boolean, String, other -> as String
                {
                    typesArray[j] = -1;
                }
            }

            sheetWritter.Write("<row r=\"");
            sheetWritter.Write(rowNum + 1 + startingRow);
            sheetWritter.Write("\">");

            for (int j = 0; j < colCnt; j++)
            {
                string stringValue = reader.GetName(j);
                sheetWritter.Write("<c r=\"");
                sheetWritter.Write(_letters[j + startingColumn]);
                sheetWritter.Write((rowNum + 1 + startingRow));
                if (stringValue.Length > 0 && (stringValue[0] == ' ' || stringValue[0] == '\t'))
                {
                    sheetWritter.Write("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                }
                else
                {
                    sheetWritter.Write("\" t=\"inlineStr\"><is><t>");
                }
                sheetWritter.Write(stringValue);
                sheetWritter.Write("</t></is></c>");
            }
            sheetWritter.Write("</row>");
            rowNum++;

            while (reader.Read())
            {
                sheetWritter.Write("<row r=\"");
                sheetWritter.Write(rowNum + 1 + startingRow);
                sheetWritter.Write("\">");

                for (int j = 0; j < colCnt; j++)
                {
                    var rawValue = reader.GetValue(j);
                    if (rawValue == null || rawValue == DBNull.Value)
                        continue;

                    if (typesArray[j] == 0 || typesArray[j] == -1)
                    {
 
                        string stringValue = rawValue.ToString()!;

                        if (stringValue.Contains('&'))
                        {
                            stringValue = stringValue.Replace("&", "&amp;");
                        }
                        if (stringValue.Contains('<'))
                        {
                            stringValue = stringValue.Replace("<", "&lt;");
                        }
                        if (stringValue.Contains('>'))
                        {
                            stringValue = stringValue.Replace(">", "&gt;");
                        }
                        if (stringValue.Contains('\"'))
                        {
                            stringValue = stringValue.Replace("\"", "&quot;");
                        }
                        if (stringValue.Contains('\''))
                        {
                            stringValue = stringValue.Replace("'", "&apos;");
                        }

                        sheetWritter.Write("<c r=\"");
                        sheetWritter.Write(_letters[j + startingColumn]);
                        sheetWritter.Write((rowNum + 1 + startingRow));
                        if (stringValue.Length > 0 && (stringValue[0] == ' ' || stringValue[0] == '\t'))
                        {
                            sheetWritter.Write("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                        }
                        else
                        {
                            sheetWritter.Write("\" t=\"inlineStr\"><is><t>");
                        }
                        sheetWritter.Write(stringValue);
                        sheetWritter.Write("</t></is></c>");
                    }
                    else if (typesArray[j] == 1)//number
                    {
                        sheetWritter.Write("<c r=\"");
                        sheetWritter.Write(_letters[j + startingColumn]);
                        sheetWritter.Write((rowNum + 1 + startingRow));
                        sheetWritter.Write("\"><v>");
                        if (rawValue is decimal decimalValue)
                        {
                            sheetWritter.Write(decimal.ToDouble(decimalValue));
                        }
                        else
                        {
                            sheetWritter.Write(rawValue);
                        }

                        sheetWritter.Write("</v></c>");
                    }
                    else if (typesArray[j] == 2) //date
                    {
                        sheetWritter.Write("<c r=\"");
                        sheetWritter.Write(_letters[j + startingColumn]);
                        sheetWritter.Write((rowNum + 1 + startingRow));
                        sheetWritter.Write($"\" s=\"{dateStyleNum}\"><v>");
                        sheetWritter.Write(((rawValue) as DateTime?)?.ToOADate());
                        sheetWritter.Write("</v></c>");

                    }
                    else if (typesArray[j] == 3) //datetime
                    {
                        if (rawValue is DateTime valDateTime && valDateTime.Year == 1000)//1000-xx-xx
                        {
                            continue;
                        }

                        sheetWritter.Write("<c r=\"");
                        sheetWritter.Write(_letters[j + startingColumn]);
                        sheetWritter.Write((rowNum + 1 + startingRow));
                        sheetWritter.Write($"\" s=\"{dateTimeStyleNum}\"><v>");
                        sheetWritter.Write(((rawValue) as DateTime?)?.ToOADate());
                        sheetWritter.Write("</v></c>");
                    }
                }

                rowNum++;
                sheetWritter.Write("</row>");
            }

            sheetWritter.Write("</sheetData>");
            sheetWritter.Write("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>");
            return rowNum - 1;
        }

    }

    public class FormattingStreamWriter : StreamWriter
    {
        private readonly IFormatProvider formatProvider;

        public FormattingStreamWriter(string path, bool append, Encoding encoding, int bufferSize, IFormatProvider formatProvider)
            : base(path, append, encoding, bufferSize)
        {
            this.formatProvider = formatProvider;
        }

        public FormattingStreamWriter(string path, bool append, Encoding encoding, IFormatProvider formatProvider)
        : base(path, append, encoding)
        {
            this.formatProvider = formatProvider;
        }
        public FormattingStreamWriter(string path, IFormatProvider formatProvider)
            : base(path)
        {
            this.formatProvider = formatProvider;
        }
        public FormattingStreamWriter(Stream stream, IFormatProvider formatProvider)
        : base(stream)
        {
            this.formatProvider = formatProvider;
        }
        public FormattingStreamWriter(Stream stream, Encoding encoding, int bufferSize, IFormatProvider formatProvider)
        : base(stream, encoding, bufferSize)
        {
            this.formatProvider = formatProvider;
        }

        public FormattingStreamWriter(Stream stream, Encoding encoding, IFormatProvider formatProvider)
        : base(stream, encoding)
        {
            this.formatProvider = formatProvider;
        }


        public override IFormatProvider FormatProvider
        {
            get
            {
                return this.formatProvider;
            }
        }
    }

    /// <summary>
    /// cellXfs
    /// </summary>
    internal class StyleInfo
    {
        public int NumFmtId;
        public int XfId;
        //public int ApplyNumberFormat;
    }

    //internal class NumberFormatInfo
    //{
    //    public string Name;
    //    public Type proposedType;
    //}

}
