using System;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using SpreadSheetTasks;

namespace SpreadSheetTasks.CsvWriter
{
    public class CsvWriter
    {
        private string _path;
        private Encoding _enconding;
        private string _rowDelimiter = "\r\n";
        private char _colDelimiter = '|';
        private bool _includeHeaders = true;

        private const char qoute = '"';
        private const char dateDelimiter = '-';
        private const char timeDelimiter = ':';
        private const char dateVsTimeDelimiter = ' ';

        public CsvWriter(string path,string rowDelimiter = "\r\n", char colDelimiter = '|', Encoding encoding = null)
        {
            if (colDelimiter == '.' || colDelimiter == dateDelimiter || colDelimiter == timeDelimiter 
                || colDelimiter == dateVsTimeDelimiter || colDelimiter == qoute)
            {
                throw new Exception($"\"{colDelimiter}\" is not supported as column separator");
            }

            if (!rowDelimiter.Contains('\n'))
            {
                throw new Exception("row delimeter have to contains \\n");
            }

            _path = path;
            _enconding = encoding?? Encoding.UTF8;
            _rowDelimiter = rowDelimiter;
            _colDelimiter = colDelimiter;
        }

      

        static CultureInfo _invariantCulture = CultureInfo.InvariantCulture;
        const int BUFFER_SIZE = 65_536;
        const int BUFFER_SIZE_HALF = BUFFER_SIZE / 2;
        char[] buffer = new char[BUFFER_SIZE];
        int currentBufferOffset = 0;
        public long Write(IDataReader datareader)
        {
            long rows = 0;
            using var fs = new StreamWriter(_path, false, _enconding);

            int fieldCount = datareader.FieldCount;

            TypeCode[] types = new TypeCode[fieldCount];
            bool[] allowNull = new bool[fieldCount];
            for (int i = 0; i < fieldCount; i++)
            {
                types[i] = Type.GetTypeCode(datareader.GetFieldType(i));
            }
            if (datareader is DbDataReader)
            {
                var schema = (datareader as IDbColumnSchemaGenerator)?.GetColumnSchema();
                for (int i = 0; i < fieldCount; i++)
                {
                    allowNull[i] = schema?[i].AllowDBNull ?? true;
                }
            }
            else
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    allowNull[i] = true;
                }
            }


            if (_includeHeaders)
            {
                for (int i = 0; i < fieldCount - 1; i++)
                {
                    fs.Write(datareader.GetName(i));
                    fs.Write(_colDelimiter);
                }
                fs.Write(datareader.GetName(fieldCount - 1));
                fs.Write(_rowDelimiter);
            }

            string tempString = "";
            int len = 0;
            while (datareader.Read())
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    if (allowNull[i] && !datareader.IsDBNull(i))
                    {
                        switch (types[i])
                        {
                            case TypeCode.Boolean:
                                bool boolVal = datareader.GetBoolean(i);
                                boolVal.TryFormat(buffer.AsSpan(currentBufferOffset), out len);
                                currentBufferOffset += len;
                                break;
                            case TypeCode.Char:
                                char valchar = datareader.GetChar(i);
                                buffer[currentBufferOffset++] = valchar;
                                break;
                            case TypeCode.Byte:
                                byte valByte = datareader.GetByte(i);
                                valByte.TryFormat(buffer.AsSpan(currentBufferOffset), out len, default, _invariantCulture);
                                currentBufferOffset += len;
                                break;
                            case TypeCode.Int16:
                                Int16 val16 = datareader.GetInt16(i);
                                val16.TryFormat(buffer.AsSpan(currentBufferOffset), out len, default, _invariantCulture);                                
                                currentBufferOffset += len;
                                break;
                            case TypeCode.Int32:
                                Int32 val = datareader.GetInt32(i);
                                val.TryFormat(buffer.AsSpan(currentBufferOffset), out len, default, _invariantCulture);
                                currentBufferOffset += len;
                                break;
                            case TypeCode.Int64:
                                Int64 val64 = datareader.GetInt64(i);
                                val64.TryFormat(buffer.AsSpan(currentBufferOffset), out len, default, _invariantCulture);
                                currentBufferOffset += len;
                                break;
                            case TypeCode.Single:
                                var valFloat = datareader.GetFloat(i);
                                valFloat.TryFormat(buffer.AsSpan(currentBufferOffset), out len, default, _invariantCulture);
                                currentBufferOffset += len;
                                break;
                            case TypeCode.Double:
                                var valDouble = datareader.GetDouble(i);
                                valDouble.TryFormat(buffer.AsSpan(currentBufferOffset), out len, default, _invariantCulture);
                                currentBufferOffset += len;
                                break;
                            case TypeCode.Decimal:
                                var valDec = datareader.GetDecimal(i);
                                valDec.TryFormat(buffer.AsSpan(currentBufferOffset), out len, default, _invariantCulture);
                                currentBufferOffset += len;
                                break;
                            case TypeCode.DateTime:
                                DateTime dtVal = datareader.GetDateTime(i);
                                //writeSimpleDateToBuffer(dtVal);
                                //writeSimpleDateTimeToBuffer(dtVal);
                                //writeDateTimeWithCulture(dtVal);
                                writeIsoDateTimeToBuffer(dtVal);
                                break;
                            case TypeCode.String:
                                tempString = datareader.GetString(i);
                                writeStringToBuffer(tempString);
                                break;
                            default:
                                tempString = datareader.GetValue(i).ToString();
                                writeStringToBuffer(tempString);
                                break;
                        }
                    }
                    if (i < fieldCount - 1)
                    {
                        buffer[currentBufferOffset++] = _colDelimiter;
                    }
                }

                for (int j = 0; j < _rowDelimiter.Length; j++)
                {
                    buffer[currentBufferOffset++] = _rowDelimiter[j];
                }

                if (currentBufferOffset >= BUFFER_SIZE_HALF)
                {
                    fs.Write(buffer, 0, currentBufferOffset);
                    currentBufferOffset = 0;
                }
                rows++;
            }

            if (currentBufferOffset > 0)
            {
                fs.Write(buffer, 0, currentBufferOffset);
                currentBufferOffset = 0;
            }

            return rows;
        }

        //private void writeDateTimeWithCulture(DateTime dtVal)
        //{
        //    if (!dtVal.TryFormat(buffer.AsSpan(currentBufferOffset), out int len, default, _invariantCulture))
        //    {
        //        throw new Exception("to small buffer!");
        //    }
        //    currentBufferOffset += len;
        //}

        // YYYY-MM-DD
        // 0123456789
        private void writeIsoDateToBuffer(DateTime dtVal)
        {
            int year = dtVal.Year;
            int month = dtVal.Month;
            int day = dtVal.Day;

            buffer[currentBufferOffset + 9] = (char)('0' + day % 10);
            buffer[currentBufferOffset + 8] = (char)('0' + day / 10);
            buffer[currentBufferOffset + 7] = dateDelimiter;

            buffer[currentBufferOffset + 6] = (char)('0' + month % 10);
            buffer[currentBufferOffset + 5] = (char)('0' + month / 10);
            buffer[currentBufferOffset + 4] = dateDelimiter;

            buffer[currentBufferOffset + 3] = (char)((year % 10) + '0');
            year /= 10;
            buffer[currentBufferOffset + 2] = (char)((year % 10) + '0');
            year /= 10;
            buffer[currentBufferOffset + 1] = (char)((year % 10) + '0');
            year /= 10;
            buffer[currentBufferOffset + 0] = (char)((year % 10) + '0');

            currentBufferOffset += 10;

            /*Span<char> span = buffer.AsSpan().Slice(currentBufferOffset, 10);

            span[9] = (char)('0' + day % 10);
            span[8] = (char)('0' + day / 10);
            span[7] = dateDelimiter;

            span[6] = (char)('0' + month % 10);
            span[5] = (char)('0' + month / 10);
            span[4] = dateDelimiter;

            span[3] = (char)((year % 10) + '0');
            year /= 10;
            span[2] = (char)((year % 10) + '0');
            year /= 10;
            span[1] = (char)((year % 10) + '0');
            year /= 10;
            span[0] = (char)((year % 10) + '0');
            
            currentBufferOffset += 10;*/
        }

        // YYYY-MM-DD HH-MM-SS
        // 0123456789012345678
        private void writeIsoDateTimeToBuffer(DateTime dtVal)
        {
            int year = dtVal.Year;
            int month = dtVal.Month;
            int day = dtVal.Day;
            int hour = dtVal.Hour;
            int minute = dtVal.Minute;
            int second = dtVal.Second;

            buffer[currentBufferOffset + 18] = (char)('0' + second % 10);
            buffer[currentBufferOffset + 17] = (char)('0' + second / 10);
            buffer[currentBufferOffset + 16] = timeDelimiter;

            buffer[currentBufferOffset + 15] = (char)('0' + minute % 10);
            buffer[currentBufferOffset + 14] = (char)('0' + minute / 10);
            buffer[currentBufferOffset + 13] = timeDelimiter;

            buffer[currentBufferOffset + 12] = (char)('0' + hour % 10);
            buffer[currentBufferOffset + 11] = (char)('0' + hour / 10);
            buffer[currentBufferOffset + 10] = dateVsTimeDelimiter;

            buffer[currentBufferOffset + 9] = (char)('0' + day % 10);
            buffer[currentBufferOffset + 8] = (char)('0' + day / 10);
            buffer[currentBufferOffset + 7] = dateDelimiter;

            buffer[currentBufferOffset + 6] = (char)('0' + month % 10);
            buffer[currentBufferOffset + 5] = (char)('0' + month / 10);
            buffer[currentBufferOffset + 4] = dateDelimiter;

            buffer[currentBufferOffset + 3] = (char)((year % 10) + '0');
            year /= 10;
            buffer[currentBufferOffset + 2] = (char)((year % 10) + '0');
            year /= 10;
            buffer[currentBufferOffset + 1] = (char)((year % 10) + '0');
            year /= 10;
            buffer[currentBufferOffset + 0] = (char)((year % 10) + '0');

            currentBufferOffset += 19;
        }

        private void writeStringToBuffer(string temp)
        {
            bool escape = false;
            int orgOffset = currentBufferOffset;

            if (temp.Length + orgOffset >= BUFFER_SIZE)
            {
                throw new Exception("buffers is too small");
            }

            for (int i = 0; i < temp.Length; i++)
            {
                char c = temp[i];
                if (c == _colDelimiter || c == '\n' || c == qoute)
                {
                    escape = true;
                    break;
                }
                buffer[currentBufferOffset++] = c;
            }
            if (!escape)
            {
                return;
            }
            else
            {
                currentBufferOffset = orgOffset;
                buffer[currentBufferOffset++] = qoute;

                for (int i = 0; i < temp.Length; i++)
                {
                    char c = temp[i];
                    buffer[currentBufferOffset++] = c;
                    if (c == qoute)
                    {
                        buffer[currentBufferOffset++] = qoute;
                    }
                }

                buffer[currentBufferOffset++] = qoute;
            } 
        }

        private string prepareStringValue(string temp)
        {
            if (temp.Contains(_colDelimiter) || temp.Contains(_rowDelimiter))
            {
                if (temp.Contains(qoute))
                {
                    // to change !! (string create/buffer)
                    temp = temp.Replace($"{qoute}", $"{qoute}{qoute}");
                }
                temp = $"{qoute}{temp}{qoute}";
            }

            return temp;
        }

    }
}
