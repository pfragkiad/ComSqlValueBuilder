using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ComSqlValueBuilder
{

    [Guid("6A629DC8-E8B3-4D4F-93DE-4A4B5AD0F416")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class SqlValueBuilder : ISqlValueBuilder
    {
        private Worksheet _sh;
        private readonly Dictionary<string, string> _fieldColumns = new Dictionary<string, string>();
        private readonly Dictionary<string, int> _fieldTypes = new Dictionary<string, int>();
        private readonly Dictionary<string, object> _values = new Dictionary<string, object>();

        public Worksheet Worksheet
        {
            get => _sh;
            set => _sh = value;
        }

        public string NullEquivalent { get; set; }

        public string TableName { get; set; }

        public string InsertSqlColumn { get; set; }

        public string UpdateSqlColumn { get; set; }

        public long FromRow { get; set; }

        public long ToRow { get; set; }

        public void AddString(string key, string column = "")
        {
            _fieldColumns[key] = column;
            _fieldTypes[key] = (int)VarType.String;
        }

        public void AddDouble(string key, string column = "")
        {
            _fieldColumns[key] = column;
            _fieldTypes[key] = (int)VarType.Double;
        }

        public void SetDoubleValue(string key, object value)
        {
            if (value == null || value.ToString() == NullEquivalent)
            {
                _values[key] = "NULL";
            }
            else
            {
                _values[key] = Convert.ToDouble(value);
            }
        }

        public void AddLong(string key, string column = "")
        {
            _fieldColumns[key] = column;
            _fieldTypes[key] = (int)VarType.Long;
        }

        public void AddBoolean(string key, string column = "")
        {
            _fieldColumns[key] = column;
            _fieldTypes[key] = (int)VarType.Boolean;
        }

        public void LoadFromSchema(Range labelCell, long firstDataRow, long lastDataRow = -1)
        {

            string sheetName = labelCell.Offset[1, 1].Value;
            _sh = labelCell.Worksheet.Parent.Sheets[sheetName] as Worksheet;

            NullEquivalent = labelCell.Offset[2, 1].Text;
            TableName = labelCell.Offset[3, 1].Text;
            InsertSqlColumn = labelCell.Offset[4, 1].Text;
            UpdateSqlColumn = labelCell.Offset[5, 1].Text;

            Range fieldsCell = labelCell.Offset[7, 1];
            long firstFieldRow = fieldsCell.Row + 1;
            long lastFieldRow = fieldsCell.End[XlDirection.xlDown].Row;

            Worksheet shSchema = labelCell.Worksheet;
            int c = labelCell.Column;
            for (long r = firstFieldRow; r <= lastFieldRow; r++)
            {
                string key = shSchema.Cells[r, c].Text;
                string column = shSchema.Cells[r, c + 1].Text;
                string sType = shSchema.Cells[r, c + 2].Text;
                switch (sType)
                {
                    case "long":
                        AddLong(key, column);
                        break;
                    case "string":
                        AddString(key, column);
                        break;
                    case "boolean":
                    case "bool":
                        AddBoolean(key, column);
                        break;
                    case "double":
                    case "float":
                        AddDouble(key, column);
                        break;
                }
            }

            FromRow = firstDataRow;

            if (lastDataRow == -1)
            {
                Range firstDataCell = _sh.Range[_fieldColumns["ID"] + firstDataRow];
                lastDataRow = GetLastRow(firstDataCell);
            }

            ToRow = lastDataRow;
        }

        public void ReadRow(long row)
        {
            _values.Clear();

            foreach (var key in _fieldColumns.Keys)
            {
                string column = _fieldColumns[key];
                if (string.IsNullOrEmpty(column)) continue;

                Range cell = _sh.Range[column + row];
                object value = cell.Value;

                if (value == null || value.ToString() == NullEquivalent || value.ToString() == "NULL")
                {
                    _values[key] = "NULL";
                }
                else if (_fieldTypes[key] == (int)VarType.String)
                {
                    _values[key] = value.ToString();
                }
                else if (_fieldTypes[key] == (int)VarType.Double)
                {
                    _values[key] = Convert.ToDouble(value);
                }
                else if (_fieldTypes[key] == (int)VarType.Long)
                {
                    _values[key] = Convert.ToInt64(value);
                }
                else if (_fieldTypes[key] == (int)VarType.Boolean)
                {
                    _values[key] = Convert.ToBoolean(value);
                }
            }
        }

        public bool IsValueNull(string key)
        {
            return _values[key].ToString() == "NULL";
        }

        public bool IsValueNotNull(string key)
        {
            return !IsValueNull(key);
        }

        public object GetValue(string key)
        {
            return _values[key];
        }

        public string GetUpdatePartialString()
        {
            if (_values.Count == 0) return "";

            List<string> values = new List<string>();
            foreach (var key in _values.Keys)
            {
                if (key == "ID") continue;

                var v = _values[key];
                if (v.ToString() == "NULL")
                    values.Add($"{key} = NULL");
                else if (v is string)
                    values.Add($"{key} = '{v}'");
                else if (v is bool vb)
                    values.Add($"{key} = {(vb ? 1 : 0)}");
                else
                    values.Add(string.Format(CultureInfo.InvariantCulture, "{0} = {1}", key, v));
            }
            return string.Join(", ", values);
        }

        public string GetInsertPartialString()
        {
            if (_values.Count == 0) return "";

            StringBuilder sql = new StringBuilder("(");
            sql.Append(string.Join(", ", _values.Keys) + ") VALUES (");

            List<string> values = new List<string>();
            foreach (var key in _values.Keys)
            {
                var v = _values[key];
                if (v.ToString() == "NULL")
                    values.Add("NULL");
                else if (v is string)
                    values.Add($"'{v}'");
                else if (v is bool vb)
                    values.Add($"{(vb ? 1 : 0)}");
                else
                    values.Add(string.Format(CultureInfo.InvariantCulture, "{0}", v));
            }
            sql.Append(string.Join(", ", values) + ")");

            return sql.ToString();
        }

        public string GetInsertSql(long row)
        {
            ReadRow(row);
            return $"INSERT INTO {TableName} {GetInsertPartialString()}";
        }

        public string GetUpdateSql(long row)
        {
            ReadRow(row);
            long id = Convert.ToInt64(GetValue("ID"));
            return $"UPDATE {TableName} SET {GetUpdatePartialString()} WHERE ID = {id}";
        }

        public void SetInsertSqls(long fromRow, long toRow, string insertSqlColumn)
        {
            for (long row = fromRow; row <= toRow; row++)
                _sh.Range[insertSqlColumn + row].Value = GetInsertSql(row);
        }

        public void SetUpdateSqls(long fromRow, long toRow, string updateSqlColumn)
        {
            for (long row = fromRow; row <= toRow; row++)
                _sh.Range[updateSqlColumn + row].Value = GetUpdateSql(row);
        }

        public void SetInsertAndUpdateSqls()
        {
            if (!string.IsNullOrEmpty(InsertSqlColumn))
                SetInsertSqls(FromRow, ToRow, InsertSqlColumn);

            if (!string.IsNullOrEmpty(UpdateSqlColumn))
                SetUpdateSqls(FromRow, ToRow, InsertSqlColumn);
        }

        private long GetLastRow(Range firstCell)
        {
            if (firstCell.Offset[1, 0].Value == null)
                return firstCell.Row;

            return firstCell.End[XlDirection.xlDown].Row;
        }
    }

    internal enum VarType
    {
        Boolean = 11,
        Double = 5,
        Long = 3,
        String = 8
    }
}