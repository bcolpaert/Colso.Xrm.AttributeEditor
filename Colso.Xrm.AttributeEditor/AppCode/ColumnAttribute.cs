using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Colso.Xrm.AttributeEditor.AppCode
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
    public sealed class ColumnAttribute : Attribute
    {
        public string Header { get; }
        public CellValues Type { get; }

        public ColumnAttribute(string header, CellValues type)
        {
            Header = header;
            Type = type;
        }
    }
}
