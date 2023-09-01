using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Colso.Xrm.AttributeEditor.AppCode
{
    public class AttributeMetadataRow
    {
        public string Action { get; set; }
        public string Entity { get; set; }

        [Column("Logical Name", CellValues.String, 200)]
        public string LogicalName { get; set; }

        [Column("Schema Name", CellValues.String, 200)]
        public string SchemaName { get; set; }

        [Column("Display Name", CellValues.String, 200)]
        public string DisplayName { get; set; }

        [Column("Type", CellValues.String)]
        public string AttributeType { get; set; }

        [Column("Description", CellValues.String, 250)]
        public string Description { get; set; }

        [Column("Field RequiredLevel", CellValues.String)]
        public string Requirement { get; set; }

        [Column("Min Value", CellValues.Number)]
        public object MinValue { get; set; }

        [Column("Max Value", CellValues.Number)]
        public object MaxValue { get; set; }

        [Column("LookupAttribute Target", CellValues.String)]
        public string LookupTarget { get; set; }

        [Column("Options (One Per Line, value:label)", CellValues.String)]
        public string Options { get; set; }

        [Column("Global Optionset Name (Name:DisplayName)", CellValues.String)]
        public string GlobalOptionsetName { get; set; }

        [Column("Precision", CellValues.Number)]
        public int? Precision { get; set; }

        [Column("Format", CellValues.String)]
        public string Format { get; set; }

        public Row ToTableRow()
        {
            var newRow = new Row();

            var properties = GetType().GetProperties();

            foreach (var column in GetColumns())
            {
                var property = properties.First(x =>
                    (x.GetCustomAttribute(typeof(ColumnAttribute)) as ColumnAttribute)?.Header == column.Header);

                newRow.AppendChild(TemplateHelper.CreateCell(column.Type, property.GetValue(this)?.ToString()));
            }

            return newRow;
        }

        public static AttributeMetadataRow FromTableRow(Row row, SharedStringTable sharedStrings)
        {
            var columns = GetColumns();
            var properties = typeof(AttributeMetadataRow).GetProperties();

            var result = new AttributeMetadataRow();

            for (var i = 0; i < columns.Length; i++)
            {
                var column = columns[i];

                var property = properties.First(x =>
                    (x.GetCustomAttribute(typeof(ColumnAttribute)) as ColumnAttribute)?.Header == column.Header);

                var stringvalue = row.GetCellValue(i, sharedStrings)?.Trim();
                property.SetValue(result, column.ConvertToRawValue(stringvalue));
            }

            return result;
        }

        public static AttributeMetadataRow FromListViewItem(ListViewItem item)
        {
            var columns = GetColumns();
            var properties = typeof(AttributeMetadataRow).GetProperties();

            var result = new AttributeMetadataRow {LogicalName = item.Text};

            for (var i = 0; i < columns.Length; i++)
            {
                var column = columns[i];

                var property = properties.First(x =>
                    (x.GetCustomAttribute(typeof(ColumnAttribute)) as ColumnAttribute)?.Header == column.Header);

                property.SetValue(result, column.ConvertToRawValue(item.SubItems[i + 1].Text));
            }

            return result;
        }

        public ListViewItem ToListViewItem()
        {
            var result = new ListViewItem();

            var properties = GetType().GetProperties();

            result.Text = Action;

            foreach (var column in GetColumns())
            {
                var property = properties.First(x =>
                    (x.GetCustomAttribute(typeof(ColumnAttribute)) as ColumnAttribute)?.Header == column.Header);

                result.SubItems.Add(property.GetValue(this)?.ToString());
            }

            return result;
        }

        public static ColumnAttribute[] GetColumns()
        {
            return typeof(AttributeMetadataRow).GetProperties()
                .Select(x => (ColumnAttribute)x.GetCustomAttribute(typeof(ColumnAttribute)))
                .Where(x => x != null)
                .ToArray();
        }

        public void UpdateFromTemplateRow(Row row, SharedStringTable sharedStrings)
        {
            if (row == null)
            {
                Action = "Delete";
                return;
            }

            Action = string.Empty;

            var columns = GetColumns();
            var properties = typeof(AttributeMetadataRow).GetProperties();

            for (var i = 0; i < columns.Length; i++)
            {
                var column = columns[i];

                var property = properties.First(x =>
                    (x.GetCustomAttribute(typeof(ColumnAttribute)) as ColumnAttribute)?.Header == column.Header);

                var oldValue = property.GetValue(this);
                var newValue = column.ConvertToRawValue(row.GetCellValue(i, sharedStrings));

                if (!Equal(oldValue, newValue))
                {
                    property.SetValue(this, newValue);
                    Action = "Update";
                }
            }
        }

        private static bool Equal(object oldValue, object newValue)
        {
            if ((oldValue == null && newValue == null) ||
                (oldValue is string && (string)oldValue == string.Empty && newValue == null))
                return true;

            return oldValue?.Equals(newValue) == true;
        }
    }
}
