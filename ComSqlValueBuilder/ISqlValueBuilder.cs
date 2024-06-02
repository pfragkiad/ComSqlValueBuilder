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
        int FromRow { get; set; }
        int ToRow { get; set; }

        void AddString(string key, string column = "");
        void AddDouble(string key, string column = "");
        void SetDoubleValue(string key, object value);
        void AddLong(string key, string column = "");
        void AddBoolean(string key, string column = "");
        void LoadFromSchema(Range labelCell, int firstDataRow=-1, int lastDataRow = -1);
        void ReadRow(int row);
        bool IsValueNull(string key);
        bool IsValueNotNull(string key);
        object GetValue(string key);
        string GetUpdatePartialString();
        string GetInsertPartialString();
        string GetInsertSql(int row);
        string GetUpdateSql(int row);
        void SetInsertSqls(int fromRow, int toRow, string insertSqlColumn);
        void SetUpdateSqls(int fromRow, int toRow, string updateSqlColumn);
        void SetInsertAndUpdateSqls();
    }
}