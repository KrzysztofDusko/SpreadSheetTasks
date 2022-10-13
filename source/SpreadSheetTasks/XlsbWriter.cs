using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace SpreadSheetTasks
{
    public sealed class XlsbWriter : ExcelWriter, IDisposable
    {
        private readonly static byte[] row25 = { 0, 25 };
        private readonly static byte[] magicArray = { 0, 0, 0, 0, 44, 1, 0, 0, 0, 1, 0, 0, 0 };
        private readonly static byte[] generalStyle = { 0, 0, 0, 0 }; // = BitConverter.GetBytes((uint)0)
        //private readonly static byte[] dateTimeStyle = { 1, 0, 0, 0 };
        //private readonly static byte[] dateStyle = { 2, 0, 0, 0 };
        //private readonly static byte[] style3 = { 3, 0, 0, 0 }; 

        private readonly static byte[] sheet1Bytes =
        {
            129,1,0,147,1,23,203,4,2,0,64,0,0,0,0,0,0,255,255,255,255,255,255,255,255,0,0,0,0,148,1,16,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,133,1,0,137,1,30,220,3,0,0,0,0,0,0,0,0,0,0,0,0,64,0,0,0,100,0,0,0,0,0,0,0,0,0,0,0,152,1,36,3,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,138,1,0,134,1,0,37,6,1,0,2,14,0,128,149,8,2,5,0,38,0,229,3,12,255,255,255,255,8,0,44,1,0,0,0,0,145,1,0,37,6,1,0,2,14,0,128,128,8,2,5,0,38,0,0,25,0,0,0,0,0,0,0,0,44,1,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,7,12,0,0,0,0,0,0,0,0,0,0,0,0,146,1,0,151,4,66,0,0,0,0,0,0,1,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,221,3,2,16,0,220,3,48,102,102,102,102,102,102,230,63,102,102,102,102,102,102,230,63,0,0,0,0,0,0,232,63,0,0,0,0,0,0,232,63,51,51,51,51,51,51,211,63,51,51,51,51,51,51,211,63,37,6,1,0,0,16,0,128,128,24,16,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,38,0,130,1,0
        };

        private readonly static byte[] stylesBin =
            {150,2,0,231,4,4,2,0,0,0,44,44,164,0,19,0,0,0,121,0,121,0,121,0,121,0,92,0,45,0,109,0,109,0,92,0,45,0,100,0,100,0,92,0,32,0,104,0,104,0,58,0,109,0,109,0,44,30,166,0,12
            ,0,0,0,121,0,121,0,121,0,121,0,92,0,45,0,109,0,109,0,92,0,45,0,100,0,100,0,232,4,0,227,4,4,1,0,0,0,43,39,220,0,0,0,144,1,0,0,0,2,0,0,7,1,0,0,0,0,0,255,2,7,0,0,0,67,0
            ,97,0,108,0,105,0,98,0,114,0,105,0,37,6,1,0,2,14,0,128,129,8,0,38,0,228,4,0,219,4,4,2,0,0,0,45,68,0,0,0,0,3,64,0,0,0,0,0,255,3,65,0,0,255,255,255,255,0,0,0,0,0,0,0,0
            ,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,45,68,17,0,0,0,3,64,0,0,0,0,0,255,3,65,0,0,255,255,255,255,0,0,0,0,0,0,0,0,0,0,0,0,0
            ,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,220,4,0,229,4,4,1,0,0,0,46,51,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0
            ,0,1,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,230,4,0,242,4,4,1,0,0,0,47,16,255,255,0,0,0,0,0,0,0,0,0,0,16,16,0,0,243,4,0,233,4,4,3,0,0,0,47,16,0,0,0,0,0,0,0,0,0,0,0,0,16
            ,16,0,0,47,16,0,0,164,0,0,0,0,0,0,0,0,0,16,16,1,0,47,16,0,0,166,0,0,0,0,0,0,0,0,0,16,16,1,0,234,4,0,235,4,4,1,0,0,0,37,6,1,0,2,17,0,128,128,24,16,0,0,0,0,0,0,0,0,0
            ,0,0,0,0,0,0,0,38,0,48,28,0,0,0,0,1,0,0,0,8,0,0,0,78,0,111,0,114,0,109,0,97,0,108,0,110,0,121,0,236,4,0,249,3,4,0,0,0,0,250,3,0,252,3,80,0,0,0,0,17,0,0,0,84,0,97,0
            ,98,0,108,0,101,0,83,0,116,0,121,0,108,0,101,0,77,0,101,0,100,0,105,0,117,0,109,0,50,0,17,0,0,0,80,0,105,0,118,0,111,0,116,0,83,0,116,0,121,0,108,0,101,0,76,0,105
            ,0,103,0,104,0,116,0,49,0,54,0,253,3,0,35,4,2,14,0,0,235,8,0,246,8,42,0,0,0,0,17,0,0,0,83,0,108,0,105,0,99,0,101,0,114,0,83,0,116,0,121,0,108,0,101,0,76,0,105,0,103
            ,0,104,0,116,0,49,0,247,8,0,236,8,0,36,0,35,4,3,15,0,0,176,16,0,178,16,50,0,0,0,0,21,0,0,0,84,0,105,0,109,0,101,0,83,0,108,0,105,0,99,0,101,0,114,0,83,0,116,0,121,0
            ,108,0,101,0,76,0,105,0,103,0,104,0,116,0,49,0,179,16,0,177,16,0,36,0,151,2,0};

        private readonly static byte[] workbookBinStart =
        {
            131,1,0,128,1,50,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2,0,0,0,120,0,108,0,1,0,0,0,55,0,1,0,0,0,54,0,5,0,0,0,50,0,52,0,51,0,50,0,54,0,153,1,12,32,0,1,0,0,0,0,0,0,0,0,0,37,6,1,0,3,15,0,128,151,16,52,24,0,0,0,67,0,58,0,92,0,115,0,113,0,108,0,115,0,92,0,84,0,101,0,115,0,116,0,121,0,90,0,97,0,112,0,105,0,115,0,117,0,88,0,108,0,115,0,98,0,92,0,38,0,37,6,1,0,0,16,0,128,129,24,130,1,0,0,0,0,0,0,0,0,47,0,0,0,49,0,51,0,95,0,110,0,99,0,114,0,58,0,49,0,95,0,123,0,49,0,54,0,53,0,48,0,56,0,68,0,54,0,57,0,45,0,67,0,70,0,56,0,55,0,45,0,52,0,55,0,54,0,57,0,45,0,56,0,52,0,53,0,54,0,45,0,68,0,52,0,65,0,52,0,48,0,49,0,49,0,51,0,49,0,53,0,54,0,55,0,125,0,47,0,0,0,47,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,38,0,135,1,0,37,6,1,0,2,16,0,128,128,24,16,0,0,0,0,13,0,0,0,255,255,255,255,0,0,0,0,38,0,158,1,29,0,0,0,0,158,22,0,0,180,105,0,0,232,38,0,0,88,2,0,0,0,0,0,0,0,0,0,0,120,136,1,0,
            143,1,0
        };

        private readonly static byte[] workbookBinEnd =
        {
            144,1,0,
            157,1,26,53,234,2,0,1,0,0,0,100,0,0,0,252,169,241,210,77,98,80,63,1,0,0,0,106,0,155,1,1,0,35,4,3,15,0,0,171,16,1,1,36,0,132,1,0
        };

        private readonly static byte[] binaryIndexBin =
        {
            42,24,0,0,0,0,32,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,149,2,0
        };

        private readonly byte[] _buffer = new byte[256];
        private readonly byte[] _buffer8 = new byte[8];

        private Stream stream;

        private int ColumnCount;
        private uint startCol;
        private uint endCol;
        private byte[] colA;
        private byte[] colZ;

        private const int rRkIntegerLowerLimit = -1 << 29;
        private const int rRkIntegerUpperLimit = (1 << 29) - 1;


        public XlsbWriter(string path)
        {
            sheetCnt = 0;
            _sstDic = new Dictionary<string, int>();
            _path = path;
            try
            {
                _newExcelFileStream = new FileStream(_path, FileMode.Create);
                _excelArchiveFile = new ZipArchive(_newExcelFileStream, ZipArchiveMode.Create);
            }
            catch (Exception)
            {
                this._path += "WARNING";
                throw new Exception("creation file error");
            }

            _sheetList = new List<(string name, string pathInArchive, string pathOnDisc, bool isHidden, string nameInArchive, int sheetId)>();
        }

        public override void AddSheet(string sheetName, bool hidden = false)
        {
            sheetCnt++;
            _sheetList.Add((sheetName, $"xl/worksheets/sheet{sheetCnt}.bin", null, hidden, $"sheet{sheetCnt}.bin", sheetCnt));
        }

        public override void WriteSheet(IDataReader dataReader, Boolean headers = true, int overLimit = -1, int startingRow = 0, int startingColumn = 0)
        {
            this.areHeaders = headers;
            _dataColReader = new DataColReader(dataReader, headers, overLimit);

            int rowNum = 0;
            ColumnCount = _dataColReader.FieldCount;

            startCol = (uint)startingColumn;
            endCol = (uint)(startCol + ColumnCount);

            colWidesArray = new double[ColumnCount];
            Array.Fill<double>(colWidesArray, -1.0);

            typesArray = new int[ColumnCount];
            newTypes = new TypeCode[ColumnCount];

            var rdr = _dataColReader.dataReader;
            for (int l = 1; l <= ColumnCount; l++)
            {
                int lenn = rdr.GetName(l - 1).Length;
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

                _dataColReader.top100.Add(arr);
                nr++;
                SetColsLengtth(ColumnCount, arr);
            }
            areNextRows = rdr.Read();
            _dataColReader.AreNextRows = areNextRows;

            if (sheetCnt != 1)
            {
                sheet1Bytes[54] = 156; // only first is selected
            }
            var newEntry = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt - 1].pathInArchive);
            stream = new BufferedStream(newEntry.Open());
            try
            {
                InitSheet();
                while (_dataColReader.Read())
                {
                    if (rowNum == 0 || areHeaders && rowNum == 1)
                    {
                        if (rowNum == 0 && areHeaders)
                        {
                            for (int i = 0; i < ColumnCount; i++)
                            {
                                typesArray[i] = 0;
                                newTypes[i] = TypeCode.String;
                            }
                        }
                        else
                        {
                            ExcelWriter.SetTypes(_dataColReader, typesArray, newTypes, ColumnCount, detectBoolenaType: true);
                        }
                    }

                    InitRow(rowNum);
                    WriteRow();

                    if (rowNum % 10000 == 0)
                    {
                        DoOn10k(rowNum);
                    }
                    rowNum++;
                }
                _rowsCount = rowNum - 1;
                stream.Write(sheet1Bytes, 218, sheet1Bytes.Length - 218); // całkowity koniec
            }
            finally
            {
                stream.Dispose();
            }
            //throw new NotImplementedException();
        }

        private void WriteRow()
        {
            for (int column = 0; column < ColumnCount; column++)
            {
                if (_dataColReader.IsDBNull(column))
                    continue;

                if (newTypes[column] == TypeCode.String) // string
                {
                    string stringValue = _dataColReader.GetString(column);
                    WriteString(stringValue, column);
                }
                else if (typesArray[column] == 5) // Memory<byte>
                {
                    var stringValue = Encoding.UTF8.GetString(((Memory<byte>)(_dataColReader.GetValue(column))).Span);
                    WriteString(stringValue, column);
                }
                else if(newTypes[column] == TypeCode.Object)
                {
                    string stringValue = _dataColReader.GetValue(column).ToString();
                    WriteString(stringValue, column);
                }
                else if (newTypes[column] == TypeCode.Boolean) // bool
                {
                    WriteBool(_dataColReader.GetBoolean(column), column);
                }
                else if (typesArray[column] == 1)//number
                {

                    switch (newTypes[column])
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

                            if (int32Value >= rRkIntegerLowerLimit && int32Value <= rRkIntegerUpperLimit)
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
                sheet1Bytes[54] = 156; // only first is selected
            }
            var newEntry = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt - 1].pathInArchive);
            stream = new BufferedStream(newEntry.Open());
            try
            {
                InitSheet();
                for (int rowNum = 0; rowNum < oneColumn.Length; rowNum++)
                {
                    string txt = oneColumn[rowNum];
                    InitRow((int)rowNum);
                    WriteString(txt, 0);
                }

                stream.Write(sheet1Bytes, 218, sheet1Bytes.Length - 218); // całkowity koniec
            }
            finally
            {
                stream.Dispose();
            }
        }

        public override void Dispose()
        {
            DoOnCompress();
            Save();
        }

        private void WriteColsWidth()
        {
            //szerokość !!!
            stream.WriteByte(134);
            stream.WriteByte(3);
            int l = 0;
            for (uint i = startCol; i < endCol; i++)
            {
                // start of column definition   
                stream.WriteByte(0);
                stream.WriteByte(60);
                stream.WriteByte(18);
                //column min
                stream.Write(BitConverter.GetBytes(i));
                // column max
                stream.Write(BitConverter.GetBytes(i));
                //width
                stream.WriteByte(0);
                stream.WriteByte((byte)(colWidesArray[l])); // .. x 7 = pixels
                stream.WriteByte(0);
                stream.WriteByte(0);

                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.WriteByte(2); // column properties /hidden etc, 2 = normal
                // end of column definition   
                l++;
            }
            stream.WriteByte(0);
            stream.WriteByte(135);
            stream.WriteByte(3);
            stream.WriteByte(0);
        }

        private void InitSheet()
        {
            colA = BitConverter.GetBytes(startCol); // start col
            colZ = BitConverter.GetBytes(endCol); // end col
            colA.CopyTo(sheet1Bytes, 40);
            colZ.CopyTo(sheet1Bytes, 44);
            stream.Write(sheet1Bytes, 0, 159); // start of file
            WriteColsWidth();
            stream.Write(sheet1Bytes, 159, 175 - 159); // BrtACBegin
            stream.WriteByte(38); // pos. 175 ?
            stream.WriteByte(0); // pos. 176 // BrtACEnd
        }

        private void InitRow(int rowNumber)
        {
            row25.CopyTo(_buffer, 0);//2 bytes
            Int32ToBuffer(rowNumber, 2);// 4 bytes -> 6
            magicArray.CopyTo(_buffer, 6); // magicArray.length = 13
            colA.CopyTo(_buffer, 6 + 13);
            colZ.CopyTo(_buffer, 6 + 13 + 4);
            stream.Write(_buffer, 0, 6 + 13 + 4 + 4);// 6 + 13 + 4 + 4 = 27
        }

        private void WriteDouble(double val, int colNum/*, int offset = 0*/, byte styleNum = 0)
        {

            _buffer[/*offset*/ +0] = 5;
            _buffer[/*offset*/ +1] = 16;//8+8
            Int32ToBuffer(colNum, /*offset*/ +2);

            //generalStyle.CopyTo(_buffer, /*offset*/ +6);
            _buffer[/*offset*/ +6] = styleNum;
            _buffer[/*offset*/ +7] = 0;
            _buffer[/*offset*/ +8] = 0;
            _buffer[/*offset*/ +9] = 0;

            BitConverter.TryWriteBytes(_buffer8, val);
            _buffer8.CopyTo(_buffer, /*offset*/ +10);
            stream.Write(_buffer, 0, 18);

            //_buffer[/*offset*/ +0] = 5;
            //_buffer[/*offset*/ +1] = 16;//8+8
            //Int32ToBuffer(colNum, /*offset*/ +2);
            //generalStyle.CopyTo(_buffer, /*offset*/ +6);
            //Int64 num3 = BitConverter.DoubleToInt64Bits(val);
            //_buffer[10] = (byte)num3;
            //_buffer[11] = (byte)(num3 >> 8);
            //_buffer[12] = (byte)(num3 >> 16);
            //_buffer[13] = (byte)(num3 >> 24);
            //_buffer[14] = (byte)(num3 >> 32);
            //_buffer[15] = (byte)(num3 >> 40);
            //_buffer[16] = (byte)(num3 >> 48);
            //_buffer[17] = (byte)(num3 >> 56);
            //stream.Write(_buffer, 0, 18);
        }

        private void WriteBool(bool val, int column)
        {
            _buffer[0] = 0x04;
            _buffer[1] = 8 + 1;
            //columnNumber
            Int32ToBuffer(column, 2);
            //styl
            generalStyle.CopyTo(_buffer, 6);
            _buffer[10] = (byte)(val ? 1 : 0); // 0 = false, 1 = true
            _buffer[11] = 1;
            stream.Write(_buffer, 0, 11);
        }

        private void WriteRkNumberInteger(int val, int colNum/*, int offset = 0*/, byte styleNum = 0)
        {

            _buffer[/*offset*/ +0] = 2;
            _buffer[/*offset*/ +1] = 12;//8+4
            Int32ToBuffer(colNum, /*offset*/ +2);
            //generalStyle.CopyTo(_buffer, /*offset*/ + 6);
            _buffer[/*offset*/ +6] = styleNum;
            _buffer[/*offset*/ +7] = 0;
            _buffer[/*offset*/ +8] = 0;
            _buffer[/*offset*/ +9] = 0;
            RkNumberIntWrite(val, /*offset*/ +10);
            stream.Write(_buffer, /*offset*/ +0, /*offset*/ +14);
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


        private void WriteString(string stringValue, int colNum)
        {
            if (_sstDic.TryGetValue(stringValue, out int index))
            {
                WriteStringFromShared(index, colNum);
            }
            else
            {
                _sstDic[stringValue] = _sstCntUnique;
                WriteStringFromShared(_sstCntUnique++, colNum);
            }
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
                _buffer[/*offset*/ +0] = 2;
                _buffer[/*offset*/ +1] = 12;//8+4
                Int32ToBuffer(colNum, /*offset*/ +2);
                //generalStyle.CopyTo(_buffer, /*offset*/ + 6);
                _buffer[/*offset*/ +6] = 2;
                _buffer[/*offset*/ +7] = 0;
                _buffer[/*offset*/ +8] = 0;
                _buffer[/*offset*/ +9] = 0;
                RkNumberGeneralWrite(d1, /*offset*/ +10, false);
                stream.Write(_buffer, /*offset*/ +0, /*offset*/ +14);
            }
        }

        private void WriteStringFromShared(int val, int colNum/*, int offset = 0*/)
        {
            //stream.WriteByte(7); // 7 = shared string
            //stream.WriteByte(8 + 4); // int  , 4 column, 4 style, 4/8 = number
            //stream.Write(BitConverter.GetBytes(colNum)); // numer kolumny
            //stream.Write(style0); // numer stylu
            //stream.Write(BitConverter.GetBytes(val)); // string number

            _buffer[/*offset*/ +0] = 7;
            _buffer[/*offset*/ +1] = 12;//8+4
            Int32ToBuffer(colNum, /*offset*/ +2);
            generalStyle.CopyTo(_buffer, /*offset*/ +6);
            Int32ToBuffer(val, /*offset*/ +10);
            stream.Write(_buffer, /*offset*/ +0, /*offset*/ +14);
        }

        internal override void FinalizeFile()
        {
            SaveSst();
            var newEntry = _excelArchiveFile.CreateEntry(@"xl/styles.bin");
            using (var str = newEntry.Open())
            {
                using var sw = new BinaryWriter(str);
                sw.Write(stylesBin);
            }

            newEntry = _excelArchiveFile.CreateEntry(@"xl/workbook.bin");
            using (var str = newEntry.Open())
            {
                using var sw = new BinaryWriter(str);
                sw.Write(workbookBinStart);

                for (int i = 0; i < _sheetList.Count; i++)
                {
                    var (name, pathInArchive, pathOnDisc, isHidden, nameInArchive, sheetId) = _sheetList[i];

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
                sw.Write(workbookBinEnd);
            }

            for (int i = 0; i < _sheetList.Count; i++)
            {
                var (_, _, _, _, _, sheetId) = _sheetList[i];
                newEntry = _excelArchiveFile.CreateEntry($@"xl/worksheets/binaryIndex{sheetId}.bin");
                using var str = newEntry.Open();
                using var sw = new BinaryWriter(str);
                sw.Write(binaryIndexBin);
            }

            newEntry = _excelArchiveFile.CreateEntry(@"[Content_Types].xml");
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
                    var (name, pathInArchive, pathOnDisc, isHidden, nameInArchive, sheetId) = _sheetList[i];

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

            newEntry = _excelArchiveFile.CreateEntry($"xl/_rels/workbook.bin.rels");
            using (var str = newEntry.Open())
            {
                using var sw = new StreamWriter(str, Encoding.UTF8);
                sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                sw.Write(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">");

                for (int i = 0; i < _sheetList.Count; i++)
                {
                    var (name, pathInArchive, pathOnDisc, isHidden, nameInArchive, sheetId) = _sheetList[i];
                    string rId = $"rId{sheetId}";
                    sw.Write($@"<Relationship Id=""{rId}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/{nameInArchive}""/>");
                }

                sw.Write($@"<Relationship Id=""rId{_sheetList.Count + 2}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.bin""/>");
                sw.Write($@"<Relationship Id=""rId{_sheetList.Count + 3}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"" Target=""sharedStrings.bin""/>");
                sw.Write(@"</Relationships>");
            }

            newEntry = _excelArchiveFile.CreateEntry($"_rels/.rels");
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
                newEntry = _excelArchiveFile.CreateEntry($"docProps/app.xml");
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
                    foreach (var (name, _, _, _, _, _) in _sheetList)
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

                newEntry = _excelArchiveFile.CreateEntry($"docProps/core.xml");
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

            foreach (var (_, _, _, _, nameInArchive, sheetId) in _sheetList)
            {
                newEntry = _excelArchiveFile.CreateEntry($"xl/worksheets/_rels/{nameInArchive}.rels");
                using var str = newEntry.Open();
                using var sw = new StreamWriter(str, Encoding.UTF8);
                sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                sw.Write(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">");
                sw.Write($@"<Relationship Id=""rId1"" Type=""http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex"" Target=""binaryIndex{sheetId}.bin""/>");
                sw.Write(@"</Relationships>");
            }

        }

        private readonly static byte[] _startSst = { 159, 1, 8 };
        private readonly static byte[] _endSst = { 160, 1, 0 };

        private void SaveSst()
        {
            var newSST = _excelArchiveFile.CreateEntry($"xl/sharedStrings.bin");
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

        private void RkNumberGeneralWrite(double d, uint offset, bool div100 = false)
        {
            // dla rk number
            // bytes[214] |=  0b00000001; = /100 flag
            // bytes[214] |=  0b00000010; = integer flag

            if (div100)
            {
                Int64 revD3 = BitConverter.DoubleToInt64Bits(100 * d);
                Int64 revD2 = revD3 >> 32;
                UInt32 revD1 = (uint)revD2;
                _buffer[offset + 0] = (byte)((revD1 % 256) | 0b00000001);
                revD1 >>= 8;
                _buffer[offset + 1] = (byte)(revD1 % 256);
                revD1 >>= 8;
                _buffer[offset + 2] = (byte)(revD1 % 256);
                revD1 >>= 8;
                _buffer[offset + 3] = (byte)(revD1 % 256);
            }
            else
            {
                Int64 revD3 = BitConverter.DoubleToInt64Bits(d);
                Int64 revD2 = revD3 >> 32;
                UInt32 revD1 = (uint)revD2;

                _buffer[offset + 0] = (byte)((revD1 % 256) & 0b11111100);
                revD1 >>= 8;
                _buffer[offset + 1] = (byte)(revD1 % 256);
                revD1 >>= 8;
                _buffer[offset + 2] = (byte)(revD1 % 256);
                revD1 >>= 8;
                _buffer[offset + 3] = (byte)(revD1 % 256);
            }
        }

        /// <summary>
        /// writes to Buffer integer in "excel mode"
        /// </summary>
        /// <param name="intNumber"></param>
        /// <param name="offset"></param>
        private void RkNumberIntWrite(Int32 intNumber, int offset)
        {
            //d <<= 2;
            //_buffer[offset] = (byte)((d % 256) | 0b00000010);
            //d >>= 8;
            //_buffer[offset + 1] = (byte)(d % 256);
            //d >>= 8;
            //_buffer[offset + 2] = (byte)(d % 256);
            //d >>= 8;
            //_buffer[offset + 3] = (byte)(d % 256);

            intNumber <<= 2;
            intNumber |= 0b00000010; // = integer flag
            //_buffer[offset++] = (byte)(d | 0b00000010);
            _buffer[offset++] = (byte)(intNumber);
            _buffer[offset++] = (byte)(intNumber >> 8);
            //d >>= 8;
            _buffer[offset++] = (byte)(intNumber >> 16);
            //d >>= 8;
            _buffer[offset++] = (byte)(intNumber >> 24);
        }

        /// <summary>
        /// integer to byte buffer
        /// </summary>
        /// <param name="intNumber"></param>
        /// <param name="offset"></param>
        private void Int32ToBuffer(Int32 intNumber, int offset)
        {
            _buffer[offset++] = (byte)intNumber;
            _buffer[offset++] = (byte)(intNumber >> 8);
            _buffer[offset++] = (byte)(intNumber >> 16);
            _buffer[offset++] = (byte)(intNumber >> 24);
        }

        private static void Int32ToSpecificBuffer(byte[] _buff, Int32 intNumber, int offset)
        {
            _buff[offset++] = (byte)intNumber;
            _buff[offset++] = (byte)(intNumber >> 8);
            _buff[offset++] = (byte)(intNumber >> 16);
            _buff[offset++] = (byte)(intNumber >> 24);
        }
    }
}
