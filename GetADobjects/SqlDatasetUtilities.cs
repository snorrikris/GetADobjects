﻿using Microsoft.SqlServer.Server;
using System;
using System.Data;
using System.Data.SqlTypes;

// Source: https://msdn.microsoft.com/en-us/library/ff878201.aspx
public static class DataSetUtilities
{
    public static void SendDataSet(DataSet ds)
    {
        if (ds == null)
        {
            throw new ArgumentException("SendDataSet requires a non-null data set.");
        }
        else
        {
            foreach (DataTable dt in ds.Tables)
            {
                SendDataTable(dt);
            }
        }
    }


    public static void SendDataTable(DataTable dt)
    {
        bool[] coerceToString;  // Do we need to coerce this column to string?
        SqlMetaData[] metaData = ExtractDataTableColumnMetaData(dt, out coerceToString);

        SqlDataRecord record = new SqlDataRecord(metaData);
        SqlPipe pipe = SqlContext.Pipe;
        pipe.SendResultsStart(record);
        try
        {
            foreach (DataRow row in dt.Rows)
            {
                for (int index = 0; index < record.FieldCount; index++)
                {
                    object value = row[index];
                    if (null != value && coerceToString[index])
                        value = value.ToString();
                    record.SetValue(index, value);
                }

                pipe.SendResultsRow(record);
            }
        }
        finally
        {
            pipe.SendResultsEnd();
        }
    }

    private static SqlMetaData[] ExtractDataTableColumnMetaData(DataTable dt, out bool[] coerceToString)
    {
        SqlMetaData[] metaDataResult = new SqlMetaData[dt.Columns.Count];
        coerceToString = new bool[dt.Columns.Count];
        for (int index = 0; index < dt.Columns.Count; index++)
        {
            DataColumn column = dt.Columns[index];
            metaDataResult[index] = SqlMetaDataFromColumn(column, out coerceToString[index]);
        }

        return metaDataResult;
    }

    private static Exception InvalidDataTypeCode(TypeCode code)
    {
        return new ArgumentException("Invalid type: " + code);
    }

    private static Exception UnknownDataType(Type clrType)
    {
        return new ArgumentException("Unknown type: " + clrType);
    }

    private static SqlMetaData SqlMetaDataFromColumn(DataColumn column, out bool coerceToString)
    {
        coerceToString = false;
        SqlMetaData sql_md = null;
        Type clrType = column.DataType;
        string name = column.ColumnName;
        switch (Type.GetTypeCode(clrType))
        {
            case TypeCode.Boolean: sql_md = new SqlMetaData(name, SqlDbType.Bit); break;
            case TypeCode.Byte: sql_md = new SqlMetaData(name, SqlDbType.TinyInt); break;
            case TypeCode.Char: sql_md = new SqlMetaData(name, SqlDbType.NVarChar, 1); break;
            case TypeCode.DateTime: sql_md = new SqlMetaData(name, SqlDbType.DateTime); break;
            case TypeCode.DBNull: throw InvalidDataTypeCode(TypeCode.DBNull);
            case TypeCode.Decimal: sql_md = new SqlMetaData(name, SqlDbType.Decimal, 18, 0); break;
            case TypeCode.Double: sql_md = new SqlMetaData(name, SqlDbType.Float); break;
            case TypeCode.Empty: throw InvalidDataTypeCode(TypeCode.Empty);
            case TypeCode.Int16: sql_md = new SqlMetaData(name, SqlDbType.SmallInt); break;
            case TypeCode.Int32: sql_md = new SqlMetaData(name, SqlDbType.Int); break;
            case TypeCode.Int64: sql_md = new SqlMetaData(name, SqlDbType.BigInt); break;
            case TypeCode.SByte: throw InvalidDataTypeCode(TypeCode.SByte);
            case TypeCode.Single: sql_md = new SqlMetaData(name, SqlDbType.Real); break;
            case TypeCode.String:
                sql_md = new SqlMetaData(name, SqlDbType.NVarChar, column.MaxLength);
                break;
            case TypeCode.UInt16: throw InvalidDataTypeCode(TypeCode.UInt16);
            case TypeCode.UInt32: throw InvalidDataTypeCode(TypeCode.UInt32);
            case TypeCode.UInt64: throw InvalidDataTypeCode(TypeCode.UInt64);
            case TypeCode.Object:
                sql_md = SqlMetaDataFromObjectColumn(name, column, clrType);
                if (sql_md == null)
                {
                    // Unknown type, try to treat it as string;
                    sql_md = new SqlMetaData(name, SqlDbType.NVarChar, column.MaxLength);
                    coerceToString = true;
                }
                break;

            default: throw UnknownDataType(clrType);
        }

        return sql_md;
    }

    private static SqlMetaData SqlMetaDataFromObjectColumn(string name, DataColumn column, Type clrType)
    {
        SqlMetaData sql_md = null;
        if (clrType == typeof(System.Byte[]) || clrType == typeof(SqlBinary) || clrType == typeof(SqlBytes) ||
    clrType == typeof(System.Char[]) || clrType == typeof(SqlString) || clrType == typeof(SqlChars))
            sql_md = new SqlMetaData(name, SqlDbType.VarBinary, column.MaxLength);
        else if (clrType == typeof(System.Guid))
            sql_md = new SqlMetaData(name, SqlDbType.UniqueIdentifier);
        else if (clrType == typeof(System.Object))
            sql_md = new SqlMetaData(name, SqlDbType.Variant);
        else if (clrType == typeof(SqlBoolean))
            sql_md = new SqlMetaData(name, SqlDbType.Bit);
        else if (clrType == typeof(SqlByte))
            sql_md = new SqlMetaData(name, SqlDbType.TinyInt);
        else if (clrType == typeof(SqlDateTime))
            sql_md = new SqlMetaData(name, SqlDbType.DateTime);
        else if (clrType == typeof(SqlDouble))
            sql_md = new SqlMetaData(name, SqlDbType.Float);
        else if (clrType == typeof(SqlGuid))
            sql_md = new SqlMetaData(name, SqlDbType.UniqueIdentifier);
        else if (clrType == typeof(SqlInt16))
            sql_md = new SqlMetaData(name, SqlDbType.SmallInt);
        else if (clrType == typeof(SqlInt32))
            sql_md = new SqlMetaData(name, SqlDbType.Int);
        else if (clrType == typeof(SqlInt64))
            sql_md = new SqlMetaData(name, SqlDbType.BigInt);
        else if (clrType == typeof(SqlMoney))
            sql_md = new SqlMetaData(name, SqlDbType.Money);
        else if (clrType == typeof(SqlDecimal))
            sql_md = new SqlMetaData(name, SqlDbType.Decimal, SqlDecimal.MaxPrecision, 0);
        else if (clrType == typeof(SqlSingle))
            sql_md = new SqlMetaData(name, SqlDbType.Real);
        else if (clrType == typeof(SqlXml))
            sql_md = new SqlMetaData(name, SqlDbType.Xml);
        else
            sql_md = null;

        return sql_md;
    }

}   // endof: class DataSetUtilities
