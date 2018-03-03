using System;
using System.Collections.Generic;
using System.Data.OleDb;
using Microsoft.Office.Interop.Access.Dao;

namespace QueryRunner.Utilities
{
    public static class TypeMapper
    {
        private static Dictionary<DataTypeEnum, OleDbType> daoTypeMap = new Dictionary<DataTypeEnum, OleDbType>
        {
            [DataTypeEnum.dbBoolean] = OleDbType.Boolean,
            [DataTypeEnum.dbByte] = OleDbType.UnsignedTinyInt,
            [DataTypeEnum.dbInteger] = OleDbType.Integer,
            [DataTypeEnum.dbLong] = OleDbType.BigInt,
            [DataTypeEnum.dbCurrency] = OleDbType.Currency,
            [DataTypeEnum.dbSingle] = OleDbType.Single,
            [DataTypeEnum.dbDouble] = OleDbType.Double,
            [DataTypeEnum.dbDate] = OleDbType.DBDate,
            [DataTypeEnum.dbBinary] = OleDbType.Binary,
            [DataTypeEnum.dbText] = OleDbType.VarChar,
            [DataTypeEnum.dbLongBinary] = OleDbType.LongVarBinary,
            [DataTypeEnum.dbMemo] = OleDbType.LongVarChar,
            [DataTypeEnum.dbGUID] = OleDbType.Guid,
            [DataTypeEnum.dbBigInt] = OleDbType.BigInt,
            [DataTypeEnum.dbVarBinary] = OleDbType.VarBinary,
            [DataTypeEnum.dbChar] = OleDbType.Char,
            [DataTypeEnum.dbNumeric] = OleDbType.Numeric,
            [DataTypeEnum.dbDecimal] = OleDbType.Decimal,
            [DataTypeEnum.dbFloat] = OleDbType.Double,
            [DataTypeEnum.dbTime] = OleDbType.DBTime,
            [DataTypeEnum.dbTimeStamp] = OleDbType.DBTimeStamp,
            [DataTypeEnum.dbAttachment] = OleDbType.Variant,
            [DataTypeEnum.dbComplexByte] = OleDbType.Variant,
            [DataTypeEnum.dbComplexInteger] = OleDbType.Variant,
            [DataTypeEnum.dbComplexLong] = OleDbType.Variant,
            [DataTypeEnum.dbComplexSingle] = OleDbType.Variant,
            [DataTypeEnum.dbComplexDouble] = OleDbType.Variant,
            [DataTypeEnum.dbComplexGUID] = OleDbType.Variant,
            [DataTypeEnum.dbComplexDecimal] = OleDbType.Variant,
            [DataTypeEnum.dbComplexText] = OleDbType.Variant
        };

        private static Dictionary<OleDbType, Type> clrTypeMap = new Dictionary<OleDbType, Type>
        {
            [OleDbType.Boolean] = typeof(bool),
            [OleDbType.UnsignedTinyInt] = typeof(byte),
            [OleDbType.Integer] = typeof(int),
            [OleDbType.BigInt] = typeof(long),
            [OleDbType.Currency] = typeof(decimal),
            [OleDbType.Single] = typeof(float),
            [OleDbType.Double] = typeof(double),
            [OleDbType.DBDate] = typeof(DateTime),
            [OleDbType.Binary] = typeof(byte),
            [OleDbType.VarChar] = typeof(string),
            [OleDbType.LongVarBinary] = typeof(byte),
            [OleDbType.LongVarChar] = typeof(string),
            [OleDbType.Guid] = typeof(Guid),
            [OleDbType.BigInt] = typeof(long),
            [OleDbType.VarBinary] = typeof(byte),
            [OleDbType.Char] = typeof(string),
            [OleDbType.Numeric] = typeof(decimal),
            [OleDbType.Decimal] = typeof(decimal),
            [OleDbType.Double] = typeof(double),
            [OleDbType.DBTime] = typeof(TimeSpan),
            [OleDbType.DBTimeStamp] = typeof(DateTime),
            [OleDbType.Variant] = typeof(object)
        };

        public static OleDbType MapDaoToOleDbType(DataTypeEnum daoType)
        {
            return daoTypeMap[daoType];
        }

        public static Type MapOleDbTypeToCLR(OleDbType oleDbType)
        {
            return clrTypeMap[oleDbType];
        }
    }
}
