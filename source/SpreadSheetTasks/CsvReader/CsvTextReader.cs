using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Text;

namespace SpreadSheetTasks.CsvReader
{
    [Obsolete]
    public class CsvTextReader : IDataReader
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

        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields | DynamicallyAccessedMemberTypes.PublicProperties)]
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
            return GetString(i);
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
            for (int i = 0; i < values.Length; i++)
            {
                values[i] = GetString(i);
            }
            return 0;
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
        private StreamReader reader;
        private readonly char[] buffer;


        private char columnDelimiter;
        private const char rowDelimiter = (char)'\n';

        private static readonly Vector256<ushort> lineVec = Vector256.Create((ushort)'\n');
        private static readonly Vector256<ushort> qouteVec = Vector256.Create((ushort)'"');
        private readonly Vector256<ushort> columnVec;


        private int NEW_LINE_LENGTH = 2;

        public bool UseIntrinsic = true;

        private void DetectDelimiter(string path)
        {
            using var fs = new StreamReader(path);
            int l = fs.Read(buffer, 0, 16_384);
            int n = buffer.AsSpan().IndexOf((char)'\n');
            if (n == -1)
            {
                n = l >= 100 ? 100 : l;
            }
            else
            {
                if (buffer[n-1] == '\r')
                {
                    NEW_LINE_LENGTH = 2;
                }
                else
                {
                    NEW_LINE_LENGTH = 1;
                }
            }
            Dictionary<char, int> dc = new()
            {
                [(char)','] = 0,
                [(char)';'] = 0,
                [(char)'|'] = 0,
                [(char)'\t'] = 0,
                [(char)':'] = 0
            };

            for (int i = 0; i < n; i++)
            {
                char b = buffer[i];
                if (dc.ContainsKey(b))
                {
                    dc[b]++;
                }
            }
            int max = 0;
            char maxDelim = '\0';
            foreach (var item in dc)
            {
                if (item.Value > max)
                {
                    max = item.Value;
                    maxDelim = item.Key;
                }
            }
            columnDelimiter = maxDelim;
            Array.Fill<char>(buffer, '\0');
        }

        public CsvTextReader(string path, Encoding encoding = null)
        {
            buffer = ArrayPool<char>.Shared.Rent(BUFFER_SIZE);

            DetectDelimiter(path);
            UseIntrinsic = Avx2.IsSupported && Bmi1.IsSupported;
            rowNumberArr = ArrayPool<int>.Shared.Rent(BUFFER_SIZE / 2);
            columnLocationsArr = ArrayPool<int>.Shared.Rent(BUFFER_SIZE);
            charBuffer = ArrayPool<char>.Shared.Rent(BUFFER_SIZE / 2);
            columnVec = Vector256.Create((ushort)this.columnDelimiter);

            Open(path, encoding);

        }

        public CsvTextReader(char colDel,string path, Encoding encoding = null)
        {
            buffer = ArrayPool<char>.Shared.Rent(BUFFER_SIZE);

            this.columnDelimiter = colDel;
            UseIntrinsic = Avx2.IsSupported && Bmi1.IsSupported;
            rowNumberArr = ArrayPool<int>.Shared.Rent(BUFFER_SIZE / 2);
            columnLocationsArr = ArrayPool<int>.Shared.Rent(BUFFER_SIZE);
            charBuffer = ArrayPool<char>.Shared.Rent(BUFFER_SIZE / 2);
            columnVec = Vector256.Create((ushort)this.columnDelimiter);

            Open(path, encoding);
        }


        private int _fieldCount = -1;
        private readonly int[] rowNumberArr;
        private readonly int[] columnLocationsArr;

        private void Open(string path, Encoding encoding = null)
        {
            if (encoding is not null)
            {
                reader = new StreamReader(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None, bufferSize: 4096/*BUFFER_SIZE*/, FileOptions.SequentialScan),encoding);
            }
            else
            {
                reader = new StreamReader(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None, bufferSize: 4096/*BUFFER_SIZE*/, FileOptions.SequentialScan));
            }
        }

        public double RelativePositionInStream()
        {
            return reader.BaseStream.Position / (double)reader.BaseStream.Length;
        }


        public void Dispose()
        {
            ArrayPool<int>.Shared.Return(rowNumberArr);
            ArrayPool<int>.Shared.Return(columnLocationsArr);
            ArrayPool<char>.Shared.Return(charBuffer);
            ArrayPool<char>.Shared.Return(buffer);

            reader.Dispose();
        }
        public void Close()
        {
            Dispose();
        }

        private int rowNumber = 1;
        private int rowNumberA = 0;
        private int rowNumberB = 0;

        //BUFFER = [rowNumberA, rowNumberA+1,...,rowNumberB-1, rowNumberB]

        int columnNumX = 0;
        const int vectorLength = 256 / 4 / 4;

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

            if (UseIntrinsic)
            {
                fixed (char* ptr2 = buffer)
                {
                    ushort* ptr = (ushort*)ptr2;
                    for (; i <= readed - vectorLength; i += vectorLength)
                    {
                        var currentDataVec = Avx2.LoadVector256(ptr + i);

                        Vector256<ushort> searchQuotes = Avx2.CompareEqual(currentDataVec, qouteVec);
                        uint quoteMask = (uint)Avx2.MoveMask(searchQuotes.AsByte());
                        int quoteOffset = quoteMask == 0 ? 32 : BitOperations.TrailingZeroCount(quoteMask);

                        Vector256<ushort> searchNewLineVec = Avx2.CompareEqual(currentDataVec, lineVec);
                        uint newLineMask = (uint)Avx2.MoveMask(searchNewLineVec.AsByte());

                        Vector256<ushort> searchColumnVec = Avx2.CompareEqual(currentDataVec, columnVec);
                        uint columnMask = (uint)Avx2.MoveMask(searchColumnVec.AsByte());

                        int colOffset = columnMask == 0 ? 32 : BitOperations.TrailingZeroCount(columnMask);
                        while (columnMask > 0 && colOffset < quoteOffset)
                        {
                            columnLocationsArr[columnNumX++] = i + (colOffset/2);
                            //*(columnLocationsPtr + columnNumX) = i + colOffset;
                            //columnNumX++;

                            //columnMask = (UInt32)(columnMask & (-(1 << colOffset) - 1));
                            columnMask = (UInt32)(columnMask - (3 << colOffset));
                            colOffset = BitOperations.TrailingZeroCount(columnMask);
                        }

                        if (newLineMask > 0)
                        {
                            int newLineOffset = BitOperations.TrailingZeroCount(newLineMask);

                            // first new Line
                            if (rowNumber == 1 && _fieldCount == -1)
                            {
                                if (buffer[i + (newLineOffset/2) - 1] == (byte)'\r')
                                {
                                    NEW_LINE_LENGTH = 2;
                                }
                                else
                                {
                                    NEW_LINE_LENGTH = 1;
                                }
                                for (_fieldCount = 0; _fieldCount <= columnNumX; _fieldCount++)
                                {
                                    if (columnLocationsArr[_fieldCount] > (i + (newLineOffset / 2)))
                                    {
                                        _fieldCount++;
                                        break;
                                    }
                                }
                                _fieldCount--;
                            }
                            while (newLineOffset < quoteOffset)//może być kilka nowych linii
                            {
                                rowNumberArr[rowNumber - rowNumberA] = i + (newLineOffset/2) + 1;
                                //*(rowNumberPtr+ rowNumber - rowNumberA) = i + newLineOffset + 1;

                                rowNumber++;

                                //newLineMask = (UInt32)(newLineMask & (-(1 << newLineOffset) - 1)); // 01010 -> 01000
                                newLineMask = (UInt32)(newLineMask - (3 << newLineOffset));

                                newLineOffset = BitOperations.TrailingZeroCount(newLineMask);
                            }
                        }

                        if (quoteMask > 0)
                        {
                            i = HandleQuotedColumns(i, quoteOffset);
                        }
                    }
                }

                rowNumberB = rowNumber;
            }

            if (readed == 0)
            {
                return false;
            }
            else if (!UseIntrinsic || readed < BUFFER_SIZE)
            {
                if (!UseIntrinsic || rowNumberArr[rowNumberB - ofsX - rowNumberA] != readed)
                {
                    char c = '\0';
                    for (; i < readed; i++)
                    {
                        c = buffer[i];

                        if (c == (byte)'\"')
                        {
                            i = HandleQuotedColumns(i, 0) + vectorLength - 1;
                        }
                        else if (c == columnDelimiter)
                        {
                            columnLocationsArr[columnNumX++] = i;

                        }
                        else if (c == rowDelimiter)
                        {
                            rowNumberArr[rowNumber - rowNumberA] = i + 1;
                            rowNumber++;
                            if (_fieldCount == -1)
                            {
                                _fieldCount = columnNumX;
                            }
                        }
                    }

                    if (readed < BUFFER_SIZE && c != (byte)'\n')
                    {
                        rowNumberArr[rowNumber - rowNumberA] = i + NEW_LINE_LENGTH;
                        rowNumber++;
                    }

                    rowNumberB = rowNumber;
                    //var endingSpan = buffer.AsSpan()[rowNumberArr[rowNumberB - 1 - rowNumberA]..readed];
                    //dodatekReczny = String.Concat(endingSpan.ToArray().Select(a => (char)a));
                }
                return !(readed < BUFFER_SIZE);
                //return false;
            }
            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private int HandleQuotedColumns(int i, int quoteOffset)
        {
            i += (quoteOffset/2 + 1);
            while (i < BUFFER_SIZE - 1)
            {
                char b1 = buffer[i + 1];
                char b0 = buffer[i];
                if (b1 != (char)'"' && b0 == (char)'"')
                {
                    break;
                }
                else if (b1 == (char)'"' && b0 == (char) '"')
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

        private readonly char[] charBuffer;

        int ofsX = 1;

        //[MethodImpl(MethodImplOptions.AggressiveInlining)]
        public ReadOnlySpan<char> GetSpan(int i)
        {
            int indx;
            ReadOnlySpan<char> sp;

            int off = 1;
            if (i < _fieldCount)
            {
                indx = columnLocationsArr[i + columnNumberNormalized] + 1;
            }
            else
            {
                indx = rowNumberArr[_recordsAffected - rowNumberA];
                off = NEW_LINE_LENGTH;
            }
            //sp = buffer.AsSpan().Slice(prevColumnIndex, indx - prevColumnIndex - off);
            sp = GenerateSpanFromBuffer(prevColumnIndex, indx - prevColumnIndex - off);
            prevColumnIndex = indx;
            return sp;
        }

        public string GetStringIgnoreQuoted(int i)
        {
            int indx;
            string s;
            int off = 1;
            if (i < _fieldCount)
            {
                indx = columnLocationsArr[i + columnNumberNormalized] + 1;
            }
            else
            {
                indx = rowNumberArr[_recordsAffected - rowNumberA];
                off = NEW_LINE_LENGTH;
            }
            s = new string(buffer, prevColumnIndex, indx - prevColumnIndex - off);
            prevColumnIndex = indx;
            return s;
        }


        public string GetString(int i)
        {
            int indx;
            string s;
            int off = 1;
            if (i < _fieldCount)
            {
                indx = columnLocationsArr[i + columnNumberNormalized] + 1 ;
            }
            else
            {
                indx = rowNumberArr[_recordsAffected - rowNumberA];
                off = NEW_LINE_LENGTH;
            }
            //s = new string(buffer,prevColumnIndex, indx - prevColumnIndex - off);
            s = GenerateStringFromBuffer(prevColumnIndex, indx - prevColumnIndex - off);
            prevColumnIndex = indx;
            return s;
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ReadOnlySpan<char> GenerateSpanFromBuffer(int start, int length)
        {
            if (buffer[start] != '"')
            {
                return buffer.AsSpan().Slice(start, length);
            }
            else
            {
                return GenerateSpanFromQuoted(start, length);
            }
        }

        //[MethodImpl(MethodImplOptions.AggressiveInlining)]
        //private unsafe ReadOnlySpan<char> generateSpanFromQuoted(int start, int length)
        //{
        //    char* tempBuff = stackalloc char[length - 2];
        //    int n = 0;
        //    int newLength = 0;
        //    while (++n < length - 1)
        //    {
        //        char c = *(bufferPtr + start + n);
        //        if (c == '"')
        //        {
        //            n++;
        //        }
        //        *(tempBuff + newLength) = c;
        //        newLength++;
        //    }
        //    return new ReadOnlySpan<char>((void *)tempBuff, newLength);
        //}

        readonly char[] charsBuff = new char[4096];
        private unsafe ReadOnlySpan<char> GenerateSpanFromQuoted(int start, int length)
        {
            int n = 0;
            int newLength = 0;
            while (++n < length - 1)
            {
                char c = buffer[start + n];
                if (c == '"')
                {
                    n++;
                }
                charsBuff[newLength] = c;
                newLength++;
            }
            return charsBuff.AsSpan()[..newLength];
        }
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private string GenerateStringFromQuoted(int start, int length)
        {
            //Span<char> tempBuff = length<1024 ? stackalloc char[length - 2] : new char[length-2];
            return GenerateSpanFromQuoted(start, length).ToString();
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private string GenerateStringFromBuffer(int start, int length)
        {
            if (buffer[start] != '"')
            {
                return new string(buffer, start, length);
            }
            else
            {
                return GenerateStringFromQuoted(start, length);
            }
        }



    }
}
