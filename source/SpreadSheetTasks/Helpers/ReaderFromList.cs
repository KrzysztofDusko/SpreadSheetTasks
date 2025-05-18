using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;

namespace SpreadSheetTasks.Helpers;

internal sealed class ReaderFromList : DbDataReader
{
    private readonly List<object?[]> _rows;
    private readonly List<string> _headers;
    private readonly int _fieldCount;
    private readonly int _rowsCnt;
    private readonly string[] _typeNames;
    private readonly Type[] _types;
    public ReaderFromList(List<object?[]> rows, List<string> headers, List<TypeCode> typeCodes)
    {
        _rows = rows;
        _headers = headers;

        _fieldCount = _headers.Count;
        _rowsCnt = _rows.Count();
        if (_rows.Count() > 0)
        {
            currentRow = _rows[0];
        }
        _typeNames = new string[_fieldCount];
        _types = new Type[_fieldCount];

        for (int i = 0; i < _fieldCount; i++)
        {
            _typeNames[i] = typeCodes[i].ToString();
            _types[i] = Type.GetType("System." + _typeNames[i]);
        }
    }

    private int currentRowNum = -1;
    private object?[] currentRow;

    public override DataTable? GetSchemaTable()
    {
        return null;
    }

    public override object this[int ordinal] => _rows[currentRowNum][ordinal];

    public override object this[string name] => throw new NotImplementedException();

    public override int Depth => throw new NotImplementedException();

    public override int FieldCount => _fieldCount;

    public override bool HasRows => _rowsCnt > 0;

    public override bool IsClosed => currentRowNum > _rowsCnt;

    public override int RecordsAffected => throw new NotImplementedException();

    public override bool GetBoolean(int ordinal)
    {
        return (bool)currentRow[ordinal];
    }

    public override byte GetByte(int ordinal)
    {
        return (byte)currentRow[ordinal];
    }

    public override long GetBytes(int ordinal, long dataOffset, byte[] buffer, int bufferOffset, int length)
    {
        throw new NotImplementedException();
    }

    public override char GetChar(int ordinal)
    {
        return (char)currentRow[ordinal];
    }

    public override long GetChars(int ordinal, long dataOffset, char[] buffer, int bufferOffset, int length)
    {
        throw new NotImplementedException();
    }

    public override string GetDataTypeName(int ordinal)
    {
        return _typeNames[ordinal];
    }

    public override DateTime GetDateTime(int ordinal)
    {
        return (DateTime)currentRow[ordinal];
    }

    public override decimal GetDecimal(int ordinal)
    {
        return (Decimal)currentRow[ordinal];
    }

    public override double GetDouble(int ordinal)
    {
        return (Double)currentRow[ordinal];
    }

    public override Type GetFieldType(int ordinal)
    {
        return _types[ordinal];
    }

    public override float GetFloat(int ordinal)
    {
        return (float)currentRow[ordinal];
    }

    public override Guid GetGuid(int ordinal)
    {
        return (Guid)currentRow[ordinal];
    }

    public override short GetInt16(int ordinal)
    {
        return (short)currentRow[ordinal];
    }

    public override int GetInt32(int ordinal)
    {
        return (int)currentRow[ordinal];
    }

    public override long GetInt64(int ordinal)
    {
        return (long)currentRow[ordinal];
    }

    public override string GetName(int ordinal)
    {
        return _headers[ordinal];
    }

    public override int GetOrdinal(string name)
    {
        throw new NotSupportedException();
    }

    public override string GetString(int ordinal)
    {
        return (string)currentRow[ordinal];
    }

    public override object GetValue(int ordinal)
    {
        return currentRow[ordinal];
    }

    public override int GetValues(object[] values)
    {
        for (int i = 0; i < _fieldCount; i++)
        {
            values[i] = currentRow[i] ?? DBNull.Value;
        }
        return _fieldCount;
    }

    public override bool IsDBNull(int ordinal)
    {
        var val = currentRow[ordinal];
        return val == null || val == DBNull.Value;
    }

    public override bool NextResult() => false;

    private int _readedRowNumber = 0;
    public override bool Read()
    {
        ++currentRowNum;
        bool res = currentRowNum < _rowsCnt;
        if (res)
        {
            currentRow = _rows[currentRowNum];
        }

        return res;
    }

    public override IEnumerator GetEnumerator()
    {
        throw new NotImplementedException();
    }
}