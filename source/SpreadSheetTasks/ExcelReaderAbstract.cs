using System;
using System.Data.Common;
using System.Globalization;
using System.Text;

namespace SpreadSheetTasks
{
    //some code from https://github.com/MarkPflug/Sylvan.Data.Excel
    //some code from https://github.com/ExcelDataReader/ExcelDataReader
    public abstract class ExcelReaderAbstract
    {
        public static CultureInfo invariantCultureInfo = CultureInfo.InvariantCulture;
        public int FieldCount { get; set; }
        public virtual int ResultsCount { get; }
        public virtual string ActualSheetName { get; set; }
        public virtual int RowCount { get => 123123123; }
        public abstract void Open(string path, bool fool1 = true, bool fool2 = false, Encoding encoding = null);

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

        public object GetValue(int i)
        {
            return innerRow[i].type switch
            {
                ExcelDataType.Null => DBNull.Value,
                ExcelDataType.Int64 => innerRow[i].int64Value,
                ExcelDataType.Double => innerRow[i].doubleValue,
                ExcelDataType.DateTime => innerRow[i].dtValue,
                ExcelDataType.Boolean => (innerRow[i].int64Value == 1),
                ExcelDataType.String => innerRow[i].strValue,
                //case ExcelDataType.Error:
                //    return "error in cell";
                _ => typeof(string),
            };
        }

        public void GetValues(object[] row)
        {
            for (int i = 0; i < row.Length; i++)
            {
                row[i] = GetValue(i);
            }
        }


        public ref FieldInfo GetNativeValue(int i)
        {
            return ref innerRow[i];
        }

        public ref FieldInfo[] GetNativeValues()
        {
            return ref innerRow;
        }

        public string GetString(int i)
        {
            return GetValue(i).ToString();
        }

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

        public Int64 GetInt64(int i)
        {
            ref var w = ref innerRow[i];
            if (w.type == ExcelDataType.Int64)
            {
                return w.int64Value;
            }
            else if (w.type == ExcelDataType.Double)
            {
                return Convert.ToInt64(w.doubleValue);
            }
            else
            {
                throw new InvalidCastException();
            }
        }

        public double GetDouble(int i)
        {
            ref var w = ref innerRow[i];
            if (w.type == ExcelDataType.Double)
            {
                return w.doubleValue;
            }
            else if (w.type == ExcelDataType.Int64)
            {
                return Convert.ToDouble(w.int64Value);
            }
            else
            {
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

    public struct FieldInfo
    {
        public ExcelDataType type;
        public string strValue;
        public Int64 int64Value;
        public double doubleValue;
        public DateTime dtValue;
        //public int xfIdx;
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
