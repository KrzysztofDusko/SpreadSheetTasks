using System;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace SpreadSheetTasks
{
    //some code from https://github.com/MarkPflug/Sylvan.Data.Excel
    //some code from https://github.com/ExcelDataReader/ExcelDataReader
    public abstract class ExcelReaderAbstract
    {
        protected static readonly CultureInfo invariantCultureInfo = CultureInfo.InvariantCulture;
        public int FieldCount { get; set; }
        public virtual int ResultsCount { get; }
        public virtual string ActualSheetName { get; set; }
        public virtual int RowCount { get => 123123123; }
        public abstract void Open(string path, bool readSharedStrings = true, bool updateMode = false, Encoding? encoding = null);
        public abstract bool Read();
        public virtual double RelativePositionInStream() => 50.0;

        /// <summary>
        /// use only after read first row  = GetValue + ToString
        /// </summary>
        /// <param name="i">column number</param>
        /// <returns></returns>
        public string GetName(int i)
        {
            return GetValue(i).ToString();
        }

        //public virtual string GetName(int i) => dbReader.GetName(i);
        public Type GetFieldType(int i)
        {
            return innerRow[i].type switch
            {
                ExcelDataType.Null => typeof(DBNull),
                ExcelDataType.Int64 => typeof(Int64),
                ExcelDataType.Double => typeof(Double),
                ExcelDataType.DateTime => typeof(DateTime),
                ExcelDataType.Boolean => typeof(bool),
                _ => typeof(string),
            };
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public object GetValue(int i)
        {
            ref var w = ref innerRow[i];
            return w.type switch
            {
                ExcelDataType.Null => DBNull.Value,
                ExcelDataType.Int32 => w.int32Value,
                ExcelDataType.Int64 => w.int64Value,
                ExcelDataType.Double => w.doubleValue,
                ExcelDataType.DateTime => w.dtValue,
                ExcelDataType.Boolean => w.boolValue,
                ExcelDataType.String => w.strValue,
                //case ExcelDataType.Error:
                //    return "error in cell";
                _ => null,
            };
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void GetValues(object[] row)
        {
            for (int i = 0; i < row.Length; i++)
            {
                ref var w = ref innerRow[i];
                row[i] = w.type switch
                {
                    ExcelDataType.Null => DBNull.Value,
                    ExcelDataType.Int32 => w.int32Value,
                    ExcelDataType.Int64 => w.int64Value,
                    ExcelDataType.Double => w.doubleValue,
                    ExcelDataType.DateTime => w.dtValue,
                    ExcelDataType.Boolean => w.boolValue,
                    ExcelDataType.String => w.strValue,
                    _ => null,
                };
            }
        }


        public ref FieldInfo GetNativeValue(int i)
        {
            return ref innerRow![i];
        }

        public ref FieldInfo[] GetNativeValues()
        {
            return ref innerRow!;
        }

        public bool TreatAllColumnsAsText { get; set; } = false;
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public virtual string GetString(int i)
        {
            ref var w = ref innerRow[i];
            if (w.type == ExcelDataType.String)
            {
                return w.strValue ?? string.Empty;
            }
            if (w.type == ExcelDataType.Null)
            {
                return string.Empty;
            }
            return w.type switch
            {
                ExcelDataType.Int32 => w.int32Value.ToString(invariantCultureInfo),
                ExcelDataType.Int64 => w.int64Value.ToString(invariantCultureInfo),
                ExcelDataType.Double => w.doubleValue.ToString(invariantCultureInfo),
                ExcelDataType.DateTime => w.dtValue.ToString(invariantCultureInfo),
                ExcelDataType.Boolean => w.boolValue ? bool.TrueString : bool.FalseString,
                _ => GetValue(i)?.ToString() ?? string.Empty,
            };
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public DateTime GetDateTime(int i)
        {
            ref var w = ref innerRow[i];
            if (w.type == ExcelDataType.DateTime)
            {
                return w.dtValue;
            }
            else
            {
                throw new InvalidCastException();
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public Int64 GetInt64(int i)
        {
            ref var w = ref innerRow[i];
            switch (w.type)
            {
                case ExcelDataType.Int64:
                    return w.int64Value;
                case ExcelDataType.Int32:
                    return w.int32Value;
                case ExcelDataType.Double:
                    return Convert.ToInt64(w.doubleValue);
                default:
                    throw new InvalidCastException();
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public Int32 GetInt32(int i)
        {
            ref var w = ref innerRow[i];
            switch (w.type)
            {
                case ExcelDataType.Int32:
                    return w.int32Value;
                case ExcelDataType.Int64:
                    return checked((int)w.int64Value);
                case ExcelDataType.Double:
                    return Convert.ToInt32(w.doubleValue);
                default:
                    throw new InvalidCastException();
            }
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public double GetDouble(int i)
        {
            ref var w = ref innerRow[i];
            switch (w.type)
            {
                case ExcelDataType.Double:
                    return w.doubleValue;
                case ExcelDataType.Int64:
                    return w.int64Value;
                case ExcelDataType.Int32:
                    return w.int32Value;
                default:
                    throw new InvalidCastException();
            }
        }

        public abstract string[] GetScheetNames();

        protected FieldInfo[] innerRow;

        public abstract void Dispose();
        //public virtual void Dispose()
        //{
        //    dbReader.Dispose();
        //}

        //Sylvan

    }

    [StructLayout(LayoutKind.Explicit)]
    public struct FieldInfo
    {
        [FieldOffset(8)]
        public ExcelDataType type;
        [FieldOffset(12)]
        public Int64 int64Value;
        [FieldOffset(12)]
        public Int32 int32Value;
        [FieldOffset(12)]
        public bool boolValue;
        [FieldOffset(12)]
        public double doubleValue;
        [FieldOffset(12)]
        public DateTime dtValue;
        [FieldOffset(0)]
        public string strValue;
    }

    public enum ExcelDataType
    {
        /// <summary>
        /// A cell that contains no value.
        /// </summary>
        Null = 0,
        /// <summary>
        /// Number
        /// </summary>
        Int32,
        Int64,
        Double,
        /// <summary>
        /// A DateTime value. This is an uncommonly used representation in .xlsx files.
        /// </summary>
        DateTime,
        /// <summary>
        /// A text field.
        /// </summary>
        String,
        /// <summary>
        /// A formula cell that contains a boolean.
        /// </summary>
        Boolean,
        /// <summary>
        /// A formula cell that contains an error.
        /// </summary>
        Error,
    }

}
