using System;
using System.Buffers;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Threading.Tasks;


namespace CsvReaderTest
{

//|             Method |     Mean |   Error |  StdDev |       Gen 0 |  Allocated |
//|------------------- |---------:|--------:|--------:|------------:|-----------:|
//|      MyGetByteSpan | 140.6 ms | 1.58 ms | 1.40 ms |           - |       5 KB |
//|  MyGetReadOnlySpan | 393.2 ms | 6.43 ms | 6.02 ms |           - |       5 KB |
//| MyGetReadOnlySpan2 | 359.5 ms | 4.09 ms | 3.63 ms |           - |       5 KB |
//|        MyGetString | 574.4 ms | 8.64 ms | 8.08 ms | 179000.0000 | 548,308 KB |
//|       SylvanString | 446.8 ms | 6.00 ms | 5.61 ms | 179000.0000 | 548,377 KB |


    public class CsvBinaryReader : IDataReader
    {
        public object this[int i] => throw new NotImplementedException();

        public object this[string name] => throw new NotImplementedException();

        public int Depth => throw new NotImplementedException();

        public bool IsClosed => throw new NotImplementedException();

        public int RecordsAffected { get => _recordsAffected; }

        public int FieldCount { get => (_fieldCount + 1); }

        public bool GetBoolean(int i)
        {
            throw new NotImplementedException();
        }

        public byte GetByte(int i)
        {
            throw new NotImplementedException();
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public char GetChar(int i)
        {
            throw new NotImplementedException();
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public IDataReader GetData(int i)
        {
            throw new NotImplementedException();
        }

        public string GetDataTypeName(int i)
        {
            throw new NotImplementedException();
        }

        public DateTime GetDateTime(int i)
        {
            throw new NotImplementedException();
        }

        public decimal GetDecimal(int i)
        {
            throw new NotImplementedException();
        }

        public double GetDouble(int i)
        {
            throw new NotImplementedException();
        }

        public Type GetFieldType(int i)
        {
            throw new NotImplementedException();
        }

        public float GetFloat(int i)
        {
            throw new NotImplementedException();
        }

        public Guid GetGuid(int i)
        {
            throw new NotImplementedException();
        }

        public short GetInt16(int i)
        {
            throw new NotImplementedException();
        }

        public int GetInt32(int i)
        {
            throw new NotImplementedException();
        }

        public long GetInt64(int i)
        {
            throw new NotImplementedException();
        }

        public string GetName(int i)
        {
            throw new NotImplementedException();
        }

        public int GetOrdinal(string name)
        {
            throw new NotImplementedException();
        }

        public DataTable GetSchemaTable()
        {
            throw new NotImplementedException();
        }

        public object GetValue(int i)
        {
            throw new NotImplementedException();
        }

        public int GetValues(object[] values)
        {
            throw new NotImplementedException();
        }

        public bool IsDBNull(int i)
        {
            throw new NotImplementedException();
        }

        public bool NextResult()
        {
            throw new NotImplementedException();
        }


        private const int BUFFER_SIZE = 65_536;
        private FileStream reader;
        private byte[] buffer;


        private byte columnDelimiter;
        private const byte rowDelimiter = (byte)'\n';

        private static readonly byte[] BOM = new byte[] { (byte)239, (byte)187, (byte)191 };

        private static readonly Vector256<byte> lineVec = Vector256.Create((byte)'\n');
        private static readonly Vector256<byte> qouteVec = Vector256.Create((byte)'"');

        private readonly Vector256<byte> columnVec;
        
        int NEW_LINE_LENGTH = 2;
        public CsvBinaryReader(string path) : this((byte)',', path)
        {

        }

        public CsvBinaryReader(byte colDel,string path)
        {
            buffer = ArrayPool<byte>.Shared.Rent(BUFFER_SIZE);
            rowNumberArr = ArrayPool<int>.Shared.Rent(BUFFER_SIZE/2); // TODO msk. liczba wierszy w "oknie".
            columnLocationsArr = ArrayPool<int>.Shared.Rent(BUFFER_SIZE); // TODO msk. liczba wkolumn w oknie
            charBuffer = ArrayPool<char>.Shared.Rent(BUFFER_SIZE / 2);

            columnVec = Vector256.Create(colDel);

            this.columnDelimiter = colDel;
            Open(path);
        }

        private int _fieldCount = -1;
        private readonly int[] rowNumberArr;
        private readonly int[] columnLocationsArr;

        private void Open(string path)
        {
            reader = new FileStream(path.ToString(), FileMode.Open, FileAccess.Read, FileShare.None, bufferSize: 4096/*BUFFER_SIZE*/, FileOptions.SequentialScan);
            HandleBom();
        }

        public void Dispose()
        {
            ArrayPool<int>.Shared.Return(rowNumberArr);
            ArrayPool<int>.Shared.Return(columnLocationsArr);
            ArrayPool<char>.Shared.Return(charBuffer);
            ArrayPool<byte>.Shared.Return(buffer);

            reader.Dispose();
        }
        public void Close()
        {
            Dispose();
        }

        public void HandleBom()
        {
            reader.Read(buffer, 0, 3);
            if (!buffer.AsSpan().StartsWith(BOM))
            {
                reader.Seek(0, SeekOrigin.Begin);
            }
        }
        private int rowNumber = 1;
        private int rowNumberA = 0;
        private int rowNumberB = 0;

        //BUFFER = [rowNumberA, rowNumberA+1,...,rowNumberB-1, rowNumberB]

        int columnNumX = 0;
        const int vectorLength = 256 / 2 / 4;

        unsafe private bool ReadBl()
        {
            int ofs = rowNumberB == 0 ? BUFFER_SIZE : rowNumberArr[rowNumberB - rowNumberA - 1]; // końcówka poprzedniego wiersza
            if (ofs < BUFFER_SIZE && buffer[ofs] == '\n')
            {
                ofs++;
            }
            if ((BUFFER_SIZE - ofs) > 0)
            {
                Array.Copy(buffer, ofs, buffer, 0, BUFFER_SIZE - ofs);
            }
            int readed = reader.Read(buffer, BUFFER_SIZE - ofs, ofs) + (BUFFER_SIZE - ofs);

            rowNumberA = rowNumberB;
            columnNumX = 0;

            int i = 0;
            fixed (byte* ptr = buffer)
            {
                for (; i <= readed - vectorLength; i += vectorLength)
                {
                    var currendDataVec = Avx2.LoadVector256(ptr + i);

                    Vector256<byte> searchQuotes = Avx2.CompareEqual(currendDataVec, qouteVec);
                    uint quoteMask = (uint)Avx2.MoveMask(searchQuotes);
                    int quoteOffset = quoteMask == 0 ? 32 : BitOperations.TrailingZeroCount(quoteMask);

                    Vector256<byte> searchNewLineVec = Avx2.CompareEqual(currendDataVec, lineVec);
                    uint newLineMask = (uint)Avx2.MoveMask(searchNewLineVec);

                    Vector256<byte> searchColumnVec = Avx2.CompareEqual(currendDataVec, columnVec);
                    uint columnMask = (uint)Avx2.MoveMask(searchColumnVec);

                    //po co wydzielać kolumny wiersze ? 

                    int colOffset = columnMask == 0 ? 32 : BitOperations.TrailingZeroCount(columnMask);
                    while (columnMask > 0 && colOffset < quoteOffset)
                    {
                        columnLocationsArr[columnNumX++] = i + colOffset;
                        //*(columnLocationsPtr + columnNumX) = i + colOffset;
                        //columnNumX++;

                        //columnMask = (UInt32)(columnMask & (-(1 << colOffset) - 1));
                        columnMask = (UInt32)(columnMask - (1 << colOffset));
                        colOffset = BitOperations.TrailingZeroCount(columnMask);
                    }

                    if (newLineMask > 0)
                    {
                        int newLineOffset = BitOperations.TrailingZeroCount(newLineMask);

                        // tutaj mamy 1. nową linię
                        if (rowNumber == 1 && _fieldCount == -1)
                        {
                            for (_fieldCount = 0; _fieldCount <= columnNumX; _fieldCount++)
                            {
                                if (columnLocationsArr[_fieldCount] >(i+newLineOffset))
                                {
                                    _fieldCount++;
                                    break;
                                }
                            }
                            _fieldCount--;
                            //row = new string[_fieldCount + 1];
                        }
                        while (newLineOffset < quoteOffset)//może być kilka nowych linii
                        {
                            rowNumberArr[rowNumber - rowNumberA] = i + newLineOffset + 1;
                            //*(rowNumberPtr+ rowNumber - rowNumberA) = i + newLineOffset + 1;

                            rowNumber++;

                            //newLineMask = (UInt32)(newLineMask & (-(1 << newLineOffset) - 1)); // 01010 -> 01000
                            newLineMask = (UInt32)(newLineMask - (1 << newLineOffset));

                            newLineOffset = BitOperations.TrailingZeroCount(newLineMask);
                        }
                    }

                    if (quoteMask > 0)
                    {
                        i = handleQuotedColumns(i, quoteOffset);
                    }
                }
            }

            rowNumberB = rowNumber;
            // TODO końówka ręcznie (< vector size)

            if (readed < BUFFER_SIZE)
            {
                if (rowNumberArr[rowNumberB - 1 - rowNumberA] != readed)
                {
                    //columnNumX = columnNumX - 8;
                    byte c = 0;
                    for (/*i = rowNumberArr[rowNumberB - 1 - rowNumberA]*/; i < readed; i++)
                    {
                        c = buffer[i];

                        if (c == (byte)'\"')
                        {
                            i = handleQuotedColumns(i, 0) + vectorLength - 1;
                        }
                        else if (c == columnDelimiter)
                        {
                            columnLocationsArr[columnNumX++] = i;
                        }
                        else if (c == rowDelimiter)
                        {
                            rowNumberArr[rowNumber - rowNumberA] = i + 1;
                            rowNumber++;
                        }
                    }

                    if (c != (byte)'\n')
                    {
                        rowNumberArr[rowNumber - rowNumberA] = i + NEW_LINE_LENGTH;
                        rowNumber++;
                    }


                    rowNumberB = rowNumber;

                    //var endingSpan = buffer.AsSpan()[rowNumberArr[rowNumberB - 1 - rowNumberA]..readed];
                    //dodatekReczny = String.Concat(endingSpan.ToArray().Select(a => (char)a));
                }
                return false;
            }

            return true;
        }



        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private int handleQuotedColumns(int i, int quoteOffset)
        {
            i += (quoteOffset + 1);
            while (i < BUFFER_SIZE - 1)
            {
                byte b1 = buffer[i + 1];
                byte b0 = buffer[i];
                if (b1 != (byte)'"' && b0 == (byte)'"')
                {
                    break;
                }
                else if (b1 == (byte)'"' && b0 == (byte)'"')
                {
                    i++;
                }
                i++;
            }
            //if (i >= buffer.Length - 1)
            //{
            //    //koniec bufora w "tekście"
            //    // = nie było końca linii = ta linia będzia miała jeszcze raz szasznę
            //    // chyhba nawet to niekonieczne..
            //    // już to poniżej = i - vectorLength + 1 być móże wystraczy ? 
            //    return i;
            //}

            i = i - vectorLength + 1;
            return i;
        }

        int _recordsAffected = 0;
        bool res = true;

        int prevColumnIndex = 0;

        int rowNumberNormalized = 0;
        int columnNumberNormalized = 0;

        public bool Read()
        {
            if (res && (_recordsAffected >= rowNumberB - 1 || rowNumberB == 0))
            {
                res = ReadBl();
                prevColumnIndex = 0;
                _recordsAffected = rowNumberA;
                if (_recordsAffected > 0)
                {
                    _recordsAffected--;
                }
            }
            ++_recordsAffected;

            if (ofsX != 0 && rowNumberA != 0)
            {
                ofsX = 0;
            }

            rowNumberNormalized = _recordsAffected - rowNumberA - ofsX;
            columnNumberNormalized = _fieldCount * rowNumberNormalized;

            return res || _recordsAffected < rowNumberB;
        }

        private char[] charBuffer;
        int ofsX = 1;

        public ReadOnlySpan<byte> GetByteSpan(int i)
        {
            if (i < _fieldCount)
            {
                int columnEndIndex = columnLocationsArr[i + columnNumberNormalized];
                var sp = buffer.AsSpan().Slice(prevColumnIndex, columnEndIndex - prevColumnIndex);
                prevColumnIndex = columnEndIndex + 1;
                return sp;
            }
            else
            {
                int indx = rowNumberArr[_recordsAffected - rowNumberA];
                var sp = buffer.AsSpan().Slice(prevColumnIndex, indx - prevColumnIndex - NEW_LINE_LENGTH);
                prevColumnIndex = indx ;
                return sp;
            }
        }

        public string GetString(int i)
        {
            return System.Text.Encoding.UTF8.GetString(GetByteSpan(i));
        }

        public ReadOnlySpan<char> GetCharSpan(int i)
        {
            if (i < _fieldCount)
            {
                int columnEndIndex = columnLocationsArr[i + columnNumberNormalized];
                int charCnt = System.Text.Encoding.UTF8.GetChars(buffer, prevColumnIndex, columnEndIndex - prevColumnIndex, charBuffer,0);
                prevColumnIndex = columnEndIndex + 1;
                return charBuffer.AsSpan().Slice(0, charCnt);
            }
            else
            {
                int indx = rowNumberArr[_recordsAffected - rowNumberA];
                int charCnt = System.Text.Encoding.UTF8.GetChars(buffer, prevColumnIndex, indx - prevColumnIndex - NEW_LINE_LENGTH, charBuffer, 0);
                prevColumnIndex = indx;
                return charBuffer.AsSpan().Slice(0, charCnt);
            }
        }

        public ReadOnlySpan<char> GetCharSpanWithBuffer(int i, Span<char> charBuff)
        {
            if (i < _fieldCount)
            {
                int columnEndIndex = columnLocationsArr[i + columnNumberNormalized];
                int charCnt = System.Text.Encoding.UTF8.GetChars(buffer.AsSpan().Slice(prevColumnIndex, columnEndIndex - prevColumnIndex), charBuff);
                prevColumnIndex = columnEndIndex + 1;
                return charBuff.Slice(0, charCnt);
            }
            else
            {
                int indx = rowNumberArr[_recordsAffected - rowNumberA];
                int charCnt = System.Text.Encoding.UTF8.GetChars(buffer.AsSpan().Slice(prevColumnIndex, indx - prevColumnIndex - NEW_LINE_LENGTH), charBuff);
                prevColumnIndex = indx;
                return charBuff.Slice(0, charCnt);
            }
        }


        //private string[] row;
        ///// <summary>
        ///// RetrieveStringRow needed !!
        ///// </summary>
        ///// <param name="i"></param>
        ///// <returns></returns>
        //public string GetStringFromRow(int i)
        //{
        //    return row[i];
        //}

        ///// <summary>
        /////  fils row string array
        ///// </summary>
        //public void RetrieveStringRow()
        //{
        //    int i = 0;
        //    for (; i < _fieldCount; i++)
        //    {
        //        int columnEndIndex = columnLocationsArr[i + columnNumberNormalized];
        //        row[i] = System.Text.Encoding.UTF8.GetString(buffer, prevColumnIndex, columnEndIndex - prevColumnIndex);

        //        prevColumnIndex = columnEndIndex + 1;
        //    }
        //    int indx = rowNumberArr[_recordsAffected - rowNumberA];
        //    row[i] = System.Text.Encoding.UTF8.GetString(buffer, prevColumnIndex, indx - prevColumnIndex);

        //    prevColumnIndex = indx;
        //}
     
    }

    internal class CsvBinaryReaderOld : IDataReader
    {
        public object this[int i] => throw new NotImplementedException();

        public object this[string name] => throw new NotImplementedException();

        public int Depth => throw new NotImplementedException();

        public bool IsClosed => throw new NotImplementedException();

        public int RecordsAffected { get => _recordsAffected; }

        public int FieldCount { get => (_fieldCount + 1); }


        public bool GetBoolean(int i)
        {
            throw new NotImplementedException();
        }

        public byte GetByte(int i)
        {
            throw new NotImplementedException();
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public char GetChar(int i)
        {
            throw new NotImplementedException();
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public IDataReader GetData(int i)
        {
            throw new NotImplementedException();
        }

        public string GetDataTypeName(int i)
        {
            throw new NotImplementedException();
        }

        public DateTime GetDateTime(int i)
        {
            throw new NotImplementedException();
        }

        public decimal GetDecimal(int i)
        {
            throw new NotImplementedException();
        }

        public double GetDouble(int i)
        {
            throw new NotImplementedException();
        }

        public Type GetFieldType(int i)
        {
            throw new NotImplementedException();
        }

        public float GetFloat(int i)
        {
            throw new NotImplementedException();
        }

        public Guid GetGuid(int i)
        {
            throw new NotImplementedException();
        }

        public short GetInt16(int i)
        {
            throw new NotImplementedException();
        }

        public int GetInt32(int i)
        {
            throw new NotImplementedException();
        }

        public long GetInt64(int i)
        {
            throw new NotImplementedException();
        }

        public string GetName(int i)
        {
            throw new NotImplementedException();
        }

        public int GetOrdinal(string name)
        {
            throw new NotImplementedException();
        }

        public DataTable GetSchemaTable()
        {
            throw new NotImplementedException();
        }

        public object GetValue(int i)
        {
            throw new NotImplementedException();
        }

        public int GetValues(object[] values)
        {
            throw new NotImplementedException();
        }

        public bool IsDBNull(int i)
        {
            throw new NotImplementedException();
        }

        public bool NextResult()
        {
            throw new NotImplementedException();
        }


        private const int BUFFER_SIZE = 65_536;
        private FileStream reader;
        private byte[] buffer;


        private byte columnDelimiter;
        private const byte rowDelimiter = (byte)'\n';

        private static readonly byte[] BOM = new byte[] { (byte)239, (byte)187, (byte)191 };

        private static readonly Vector256<byte> lineVec = Vector256.Create((byte)'\n');
        private static readonly Vector256<byte> qouteVec = Vector256.Create((byte)'"');

        private readonly Vector256<byte> columnVec;

        int NEW_LINE_LENGTH = 2;
        public CsvBinaryReaderOld(string path) : this((byte)',', path)
        {

        }

        public CsvBinaryReaderOld(byte colDel, string path)
        {
            buffer = ArrayPool<byte>.Shared.Rent(BUFFER_SIZE);
            rowNumberArr = ArrayPool<int>.Shared.Rent(BUFFER_SIZE / 2); // TODO msk. liczba wierszy w "oknie".
            columnLocationsArr = ArrayPool<int>.Shared.Rent(BUFFER_SIZE); // TODO msk. liczba wkolumn w oknie
            charBuffer = ArrayPool<char>.Shared.Rent(BUFFER_SIZE / 2);

            columnVec = Vector256.Create(colDel);

            this.columnDelimiter = colDel;
            Open(path);
        }

        private int _fieldCount = 0;
        private readonly int[] rowNumberArr;
        private readonly int[] columnLocationsArr;

        private void Open(string path)
        {
            reader = new FileStream(path.ToString(), FileMode.Open, FileAccess.Read, FileShare.None, bufferSize: 4096/*BUFFER_SIZE*/, FileOptions.SequentialScan);
            HandleBom();
        }

        public void Dispose()
        {
            ArrayPool<int>.Shared.Return(rowNumberArr);
            ArrayPool<int>.Shared.Return(columnLocationsArr);
            ArrayPool<char>.Shared.Return(charBuffer);
            ArrayPool<byte>.Shared.Return(buffer);

            reader.Dispose();
        }
        public void Close()
        {
            Dispose();
        }

        public void HandleBom()
        {
            reader.Read(buffer, 0, 3);
            if (!buffer.AsSpan().StartsWith(BOM))
            {
                reader.Seek(0, SeekOrigin.Begin);
            }
        }
        private int rowNumber = 1;
        private int rowNumberA = 0;
        private int rowNumberB = 0;

        //BUFFER = [rowNumberA, rowNumberA+1,...,rowNumberB-1, rowNumberB]

        int columnNumX = 0;
        const int vectorLength = 256 / 2 / 4;



        unsafe private bool ReadBl()
        {
            int ofs = rowNumberB == 0 ? BUFFER_SIZE : rowNumberArr[rowNumberB - rowNumberA - 1]; // końcówka poprzedniego wiersza
            if (ofs < BUFFER_SIZE && buffer[ofs] == '\n')
            {
                ofs++;
            }
            if ((BUFFER_SIZE - ofs) > 0)
            {
                Array.Copy(buffer, ofs, buffer, 0, BUFFER_SIZE - ofs);
            }
            int readed = reader.Read(buffer, BUFFER_SIZE - ofs, ofs) + (BUFFER_SIZE - ofs);

            rowNumberA = rowNumberB;
            columnNumX = 0;

            int i = 0;
            fixed (byte* ptr = buffer)
            {
                for (; i <= readed - vectorLength; i += vectorLength)
                {
                    var currendDataVec = Avx2.LoadVector256(ptr + i);

                    Vector256<byte> searchQuotes = Avx2.CompareEqual(currendDataVec, qouteVec);
                    uint quoteMask = (uint)Avx2.MoveMask(searchQuotes);
                    int quoteOffset = quoteMask == 0 ? 32 : BitOperations.TrailingZeroCount(quoteMask);

                    Vector256<byte> searchNewLineVec = Avx2.CompareEqual(currendDataVec, lineVec);
                    uint newLineMask = (uint)Avx2.MoveMask(searchNewLineVec);

                    Vector256<byte> searchColumnVec = Avx2.CompareEqual(currendDataVec, columnVec);
                    uint columnMask = (uint)Avx2.MoveMask(searchColumnVec);
                    int colOffset = columnMask == 0 ? 32 : BitOperations.TrailingZeroCount(columnMask);

                    //po co wydzielać kolumny wiersze ? 
                    if (columnMask > 0 && newLineMask == 0)
                    {
                        while (columnMask > 0 && colOffset < quoteOffset)
                        {
                            columnLocationsArr[columnNumX++] = i + colOffset;
                            //*(columnLocationsPtr + columnNumX) = i + colOffset;
                            //columnNumX++;

                            //columnMask = (UInt32)(columnMask & (-(1 << colOffset) - 1));
                            columnMask = (UInt32)(columnMask - (1 << colOffset));

                            colOffset = BitOperations.TrailingZeroCount(columnMask);
                        }
                    }
                    else if (newLineMask > 0)
                    {
                        int newLineOffset = BitOperations.TrailingZeroCount(newLineMask);

                        //kolumny przed nową linią
                        while (columnMask > 0 && colOffset < newLineOffset && colOffset < quoteOffset)
                        {
                            columnLocationsArr[columnNumX++] = i + colOffset;
                            //*(columnLocationsPtr + columnNumX) = i + colOffset;
                            //columnNumX++;

                            //columnMask = (UInt32) (columnMask & (-(1 << colOffset) - 1));
                            columnMask = (UInt32)(columnMask - (1 << colOffset));

                            colOffset = BitOperations.TrailingZeroCount(columnMask);
                        }

                        // tutaj mamy 1. nową linię
                        if (rowNumber == 1)
                        {
                            _fieldCount = columnNumX;
                            //row = new string[_fieldCount + 1];
                        }
                        while (newLineOffset < quoteOffset)//może być kilka nowych linii
                        {
                            rowNumberArr[rowNumber - rowNumberA] = i + newLineOffset + 1;
                            //*(rowNumberPtr+ rowNumber - rowNumberA) = i + newLineOffset + 1;

                            rowNumber++;

                            //newLineMask = (UInt32)(newLineMask & (-(1 << newLineOffset) - 1)); // 01010 -> 01000
                            newLineMask = (UInt32)(newLineMask - (1 << newLineOffset));

                            newLineOffset = BitOperations.TrailingZeroCount(newLineMask);

                            // kolumny po nowej linii ale przed cudzysłowiem
                            while (columnMask > 0 && colOffset < quoteOffset && colOffset < newLineOffset)
                            {
                                columnLocationsArr[columnNumX++] = i + colOffset;
                                //*(columnLocationsPtr + columnNumX) = i + colOffset;
                                //columnNumX++;

                                //columnMask = (UInt32)(columnMask & (-(1 << colOffset) - 1));
                                columnMask = (UInt32)(columnMask - (1 << colOffset));
                                colOffset = BitOperations.TrailingZeroCount(columnMask);
                            }
                        }
                    }

                    if (quoteMask > 0)
                    {
                        i = handleQuotedColumns(i, quoteOffset);
                    }
                }
            }

            rowNumberB = rowNumber;
            // TODO końówka ręcznie (< vector size)

            if (readed < BUFFER_SIZE)
            {
                if (rowNumberArr[rowNumberB - 1 - rowNumberA] != readed)
                {
                    //columnNumX = columnNumX - 8;
                    byte c = 0;
                    for (/*i = rowNumberArr[rowNumberB - 1 - rowNumberA]*/; i < readed; i++)
                    {
                        c = buffer[i];

                        if (c == (byte)'\"')
                        {
                            i = handleQuotedColumns(i, 0) + vectorLength - 1;
                        }
                        else if (c == columnDelimiter)
                        {
                            columnLocationsArr[columnNumX++] = i;
                        }
                        else if (c == rowDelimiter)
                        {
                            rowNumberArr[rowNumber - rowNumberA] = i + 1;
                            rowNumber++;
                        }
                    }

                    if (c != (byte)'\n')
                    {
                        rowNumberArr[rowNumber - rowNumberA] = i + NEW_LINE_LENGTH;
                        rowNumber++;
                    }


                    rowNumberB = rowNumber;

                    //var endingSpan = buffer.AsSpan()[rowNumberArr[rowNumberB - 1 - rowNumberA]..readed];
                    //dodatekReczny = String.Concat(endingSpan.ToArray().Select(a => (char)a));
                }
                return false;
            }

            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private int handleQuotedColumns(int i, int quoteOffset)
        {
            i += (quoteOffset + 1);
            while (i < BUFFER_SIZE - 1)
            {
                byte b1 = buffer[i + 1];
                byte b0 = buffer[i];
                if (b1 != (byte)'"' && b0 == (byte)'"')
                {
                    break;
                }
                else if (b1 == (byte)'"' && b0 == (byte)'"')
                {
                    i++;
                }
                i++;
            }
            //if (i >= buffer.Length - 1)
            //{
            //    //koniec bufora w "tekście"
            //    // = nie było końca linii = ta linia będzia miała jeszcze raz szasznę
            //    // chyhba nawet to niekonieczne..
            //    // już to poniżej = i - vectorLength + 1 być móże wystraczy ? 
            //    return i;
            //}

            i = i - vectorLength + 1;
            return i;
        }

        int _recordsAffected = 0;
        bool res = true;

        int prevColumnIndex = 0;

        int rowNumberNormalized = 0;
        int columnNumberNormalized = 0;

        public bool Read()
        {
            if (res && (_recordsAffected >= rowNumberB - 1 || rowNumberB == 0))
            {
                res = ReadBl();
                prevColumnIndex = 0;
                _recordsAffected = rowNumberA;
                if (_recordsAffected > 0)
                {
                    _recordsAffected--;
                }
            }
            ++_recordsAffected;

            if (ofsX != 0 && rowNumberA != 0)
            {
                ofsX = 0;
            }

            rowNumberNormalized = _recordsAffected - rowNumberA - ofsX;
            columnNumberNormalized = _fieldCount * rowNumberNormalized;

            return res || _recordsAffected < rowNumberB;
        }

        private char[] charBuffer;
        int ofsX = 1;


        public ReadOnlySpan<byte> GetByteSpan(int i)
        {
            if (i < _fieldCount)
            {
                int columnEndIndex = columnLocationsArr[i + columnNumberNormalized];
                var sp = buffer.AsSpan().Slice(prevColumnIndex, columnEndIndex - prevColumnIndex);
                prevColumnIndex = columnEndIndex + 1;
                return sp;
            }
            else
            {
                int indx = rowNumberArr[_recordsAffected - rowNumberA];
                var sp = buffer.AsSpan().Slice(prevColumnIndex, indx - prevColumnIndex - NEW_LINE_LENGTH);
                prevColumnIndex = indx;
                return sp;
            }
        }

        public string GetString(int i)
        {
            return System.Text.Encoding.UTF8.GetString(GetByteSpan(i));
        }

        public ReadOnlySpan<char> GetCharSpan(int i)
        {
            if (i < _fieldCount)
            {
                int columnEndIndex = columnLocationsArr[i + columnNumberNormalized];
                int charCnt = System.Text.Encoding.UTF8.GetChars(buffer, prevColumnIndex, columnEndIndex - prevColumnIndex, charBuffer, 0);
                prevColumnIndex = columnEndIndex + 1;
                return charBuffer.AsSpan().Slice(0, charCnt);
            }
            else
            {
                int indx = rowNumberArr[_recordsAffected - rowNumberA];
                int charCnt = System.Text.Encoding.UTF8.GetChars(buffer, prevColumnIndex, indx - prevColumnIndex - NEW_LINE_LENGTH, charBuffer, 0);
                prevColumnIndex = indx;
                return charBuffer.AsSpan().Slice(0, charCnt);
            }
        }

        public ReadOnlySpan<char> GetCharSpanWithBuffer(int i, Span<char> charBuff)
        {
            if (i < _fieldCount)
            {
                int columnEndIndex = columnLocationsArr[i + columnNumberNormalized];
                int charCnt = System.Text.Encoding.UTF8.GetChars(buffer.AsSpan().Slice(prevColumnIndex, columnEndIndex - prevColumnIndex), charBuff);
                prevColumnIndex = columnEndIndex + 1;
                return charBuff.Slice(0, charCnt);
            }
            else
            {
                int indx = rowNumberArr[_recordsAffected - rowNumberA];
                int charCnt = System.Text.Encoding.UTF8.GetChars(buffer.AsSpan().Slice(prevColumnIndex, indx - prevColumnIndex - NEW_LINE_LENGTH), charBuff);
                prevColumnIndex = indx;
                return charBuff.Slice(0, charCnt);
            }
        }


        //private string[] row;
        ///// <summary>
        ///// RetrieveStringRow needed !!
        ///// </summary>
        ///// <param name="i"></param>
        ///// <returns></returns>
        //public string GetStringFromRow(int i)
        //{
        //    return row[i];
        //}

        ///// <summary>
        /////  fils row string array
        ///// </summary>
        //public void RetrieveStringRow()
        //{
        //    int i = 0;
        //    for (; i < _fieldCount; i++)
        //    {
        //        int columnEndIndex = columnLocationsArr[i + columnNumberNormalized];
        //        row[i] = System.Text.Encoding.UTF8.GetString(buffer, prevColumnIndex, columnEndIndex - prevColumnIndex);

        //        prevColumnIndex = columnEndIndex + 1;
        //    }
        //    int indx = rowNumberArr[_recordsAffected - rowNumberA];
        //    row[i] = System.Text.Encoding.UTF8.GetString(buffer, prevColumnIndex, indx - prevColumnIndex);

        //    prevColumnIndex = indx;
        //}

    }
}
