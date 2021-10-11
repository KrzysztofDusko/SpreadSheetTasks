using System;
using System.Data.Common;

namespace SpreadSheetTasks
{
    public abstract class ExcelReaderAbstract
    {
        protected DbDataReader dbReader = null;
        protected string sylvanFilePath = null;
        public int FieldCount { get; set; }
        public virtual int ResultsCount { get; }
        public virtual string ActualSheetName { get; set; }
        public virtual int RowCount { get => 123123123; }
        public abstract void Open(string path, bool fool1 = true, bool fool2 = false);
        public virtual bool Read()
        {
            Array.Clear(innerRow, 0, FieldCount);
            return dbReader.Read();
        }

        public virtual string GetName(int i) => dbReader.GetName(i);
        public virtual Type GetFieldType(int i)
        {
            string rawVal = dbReader.GetValue(i).ToString();
            if (long.TryParse(rawVal, out long longValue))
            {
                innerRow[i] = longValue;
                return typeof(long);
            }
            else if (double.TryParse(rawVal, out double doubleValue))
            {
                innerRow[i] = doubleValue;
                return typeof(double);
            }
            else if (DateTime.TryParse(rawVal, out DateTime dateTimeValue))
            {
                innerRow[i] = dateTimeValue;
                return typeof(DateTime);
            }
            innerRow[i] = rawVal;
            return typeof(string);
        }

        public virtual object GetValue(int i)
        {
            if (innerRow[i] == null)
            {
                return dbReader.GetValue(i);
            }
            return innerRow[i];
        }

        public virtual void GetValues(object[] row)
        {
            for (int i = 0; i < row.Length; i++)
            {
                row[i] = GetValue(i);
            }
        }

        public string GetString(int i)
        {
            return GetValue(i).ToString();
        }

        public DateTime GetDateTime(int i)
        {
            var val = GetValue(i);
            return (DateTime)val;
        }

        public Int64 GetInt64(int i)
        {
            var val = GetValue(i);
            return (Int64) val;
        }

        public double GetDouble(int i)
        {
            var val = GetValue(i);
            return (double) val;
        }

        public T GetQueryString<T>(int i)
        {
            var val = GetValue(i);
            return (T)val;
        }

        public abstract string[] GetScheetNames();

        protected object[] innerRow;

        public virtual void Dispose()
        {
            dbReader.Dispose();
        }

    }
}
