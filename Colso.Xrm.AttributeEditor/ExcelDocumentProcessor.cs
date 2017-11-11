using System.Collections.Generic;
using System.IO;
using System.Linq;
using Colso.Xrm.AttributeEditor.AppCode;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Colso.Xrm.AttributeEditor
{
    public interface IExcelDocumentProcessor
    {
        ExcelDocumentProcessor.WorkbookSheet Read(string filename, string sheetName);
        void SaveWorkbook(string filename, string sheetName, ExcelDocumentProcessor.WorkbookSheet sheet);
    }

    public class ExcelDocumentProcessor : IExcelDocumentProcessor
    {
        public WorkbookSheet Read(string filename, string sheetName)
        {
            using (FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read,
                            FileShare.ReadWrite))
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
            {
                var sheet = document.WorkbookPart.Workbook.Sheets.Cast<Sheet>()
                    .Where(s => s.Name == sheetName).FirstOrDefault();

                if (sheet == null)
                    return null;

                var sheetid = sheet.Id;
                var sheetData = ((WorksheetPart)document.WorkbookPart.GetPartById(sheetid)).Worksheet
                    .GetFirstChild<SheetData>();

                var templateRows = sheetData.ChildElements.Cast<Row>().ToArray();
                // For shared strings, look up the value in the shared strings table.
                var stringTable = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>()
                    .FirstOrDefault()?.SharedStringTable;

                var rowResult = new List<WorkbookRow>();

                foreach (var row in templateRows)
                {
                    var columnsResult = new List<WorkbookCell>();

                    foreach (var cell in row.ChildElements.Cast<Cell>())
                    {
                        string cellText;

                        if (cell.DataType != null && stringTable != null
                            && cell.DataType.HasValue && cell.DataType == CellValues.SharedString
                            && int.Parse(cell.CellValue.InnerText) < stringTable.ChildElements.Count)
                        {
                            cellText = stringTable.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText;
                        }
                        else
                        {
                            cellText = cell.CellValue?.InnerText;
                        }

                        columnsResult.Add(new WorkbookCell
                        {
                            Value = cellText
                        });
                    }

                    rowResult.Add(new WorkbookRow
                    {
                        Cells = columnsResult
                    });
                }

                return new WorkbookSheet
                {
                    Rows = rowResult
                };
            }
        }

        public void SaveWorkbook(string filename, string sheetName, WorkbookSheet sheet)
        {
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                TemplateHelper.InitDocument(document, out var workbookPart, out var sheets);
                TemplateHelper.CreateSheet(workbookPart, sheets, sheetName, out var sheetData);

                foreach (var row in sheet.Rows)
                {
                    var workbookRow = new Row();

                    foreach (var cell in row.Cells)
                        workbookRow.AppendChild(TemplateHelper.CreateCell((CellValues)cell.Type, cell.Value));

                    sheetData.AppendChild(workbookRow);
                }

                workbookPart.Workbook.Save();
            }
        }

        public class WorkbookSheet
        {
            public IEnumerable<WorkbookRow> Rows { get; set; }
        }

        public class WorkbookRow
        {
            public IEnumerable<WorkbookCell> Cells { get; set; }
        }

        public class WorkbookCell
        {
            public string Value { get; set; }
            public CellTypes Type { get; set; }
        }

        public enum CellTypes
        {
            Boolean = 0,
            Number = 1,
            Error = 2,
            SharedString = 3,
            String = 4,
            InlineString = 5,
            Date = 6
        }
    }
}
