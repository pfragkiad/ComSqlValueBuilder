using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ComSqlValueBuilder
{

    [Guid("D8B9A2F1-C7E5-4F6B-8F0E-FA4F9A22B1C4")]
    public interface ISqlValueBuilder
    {
        string NullEquivalent { get; set; }
        Worksheet Worksheet { get; set; }
        string TableName { get; set; }
        string InsertSqlColumn { get; set; }
        string UpdateSqlColumn { get; set; }
        long FromRow { get; set; }
        long ToRow { get; set; }

        void AddString(string key, string column = "");
        void AddDouble(string key, string column = "");
        void SetDoubleValue(string key, object value);
        void AddLong(string key, string column = "");
        void AddBoolean(string key, string column = "");
        void LoadFromSchema(Range labelCell, long firstDataRow, long lastDataRow = -1);
        void ReadRow(long row);
        bool IsValueNull(string key);
        bool IsValueNotNull(string key);
        object GetValue(string key);
        string GetUpdatePartialString();
        string GetInsertPartialString();
        string GetInsertSql(long row);
        string GetUpdateSql(long row);
        void SetInsertSqls(long fromRow, long toRow, string insertSqlColumn);
        void SetUpdateSqls(long fromRow, long toRow, string updateSqlColumn);
        void SetInsertAndUpdateSqls();
    }
}