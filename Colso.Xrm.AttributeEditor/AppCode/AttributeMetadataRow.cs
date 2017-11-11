using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Colso.Xrm.AttributeEditor.AppCode
{
    public class AttributeMetadataRow
    {
        public string Action { get; set; }
        public string Entity { get; set; }

        [Column("Logical Name", CellValues.String)]
        public string LogicalName { get; set; }

        [Column("Schema Name", CellValues.String)]
        public string SchemaName { get; set; }

        [Column("Display Name", CellValues.String)]
        public string DisplayName { get; set; }

        [Column("Type", CellValues.String)]
        public string AttributeType { get; set; }

        [Column("Field RequiredLevel", CellValues.String)]
        public string Requirement { get; set; }

        [Column("LookupAttribute Target", CellValues.String)]
        public string LookupTarget { get; set; }

        [Column("Options (One Per Line, value:label)", CellValues.String)]
        public string Options { get; set; }

        [Column("Global Optionset Name (Name:DisplayName)", CellValues.String)]
        public string GlobalOptionsetName { get; set; }

        [Column("Precision", CellValues.String)]
        public string Precision { get; set; }

        [Column("Date Format", CellValues.String)]
        public string DateFormat { get; set; }

        public ExcelDocumentProcessor.WorkbookRow ToTableRow()
        {
            var cells = new List<ExcelDocumentProcessor.WorkbookCell>();

            var properties = GetType().GetProperties();

            foreach (var column in GetColumns())
            {
                var property = properties.First(x =>
                    (x.GetCustomAttribute(typeof(ColumnAttribute)) as ColumnAttribute)?.Header == column.Header);

                cells.Add(new ExcelDocumentProcessor.WorkbookCell
                {
                    Type = (ExcelDocumentProcessor.CellTypes)column.Type,
                    Value = property.GetValue(this) as string
                });
            }

            return new ExcelDocumentProcessor.WorkbookRow
            {
                Cells = cells
            };
        }

        public static AttributeMetadataRow FromTableRow(ExcelDocumentProcessor.WorkbookRow row)
        {
            var columns = GetColumns();
            var properties = typeof(AttributeMetadataRow).GetProperties();

            var result = new AttributeMetadataRow();

            for (var i = 0; i < columns.Length; i++)
            {
                var column = columns[i];

                var property = properties.First(x =>
                    (x.GetCustomAttribute(typeof(ColumnAttribute)) as ColumnAttribute)?.Header == column.Header);

                var value = row.Cells.Count() > i ? row.Cells.Skip(i).First().Value.Trim() : null;

                property.SetValue(result, value);
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

                property.SetValue(result, item.SubItems[i+1].Text);
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

                result.SubItems.Add((string)property.GetValue(this));
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

        public void UpdateFromTemplateRow(ExcelDocumentProcessor.WorkbookRow row)
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
                var newValue = row.Cells.Count() > i ? row.Cells.Skip(i).First().Value : null;

                if (!Equal(oldValue, newValue))
                {
                    property.SetValue(this, newValue);
                    Action = "Update";
                }
            }
        }

        private bool Equal(object oldValue, string newValue)
        {
            if (oldValue == null && newValue == null || oldValue == null && newValue == string.Empty ||
                oldValue is string && (string)oldValue == string.Empty && newValue == null)
                return true;

            return oldValue.Equals(newValue);
        }
    }
}
