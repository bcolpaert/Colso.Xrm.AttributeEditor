using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Colso.Xrm.AttributeEditor.AppCode
{
    public class TemplateHelper
    {

        public static void InitDocument(SpreadsheetDocument document, out WorkbookPart workbookPart, out Sheets sheets)
        {
            workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            sheets = workbookPart.Workbook.AppendChild(new Sheets());
        }

        public static void CreateSheet(WorkbookPart workbookPart, Sheets sheets, string name, out SheetData sheetData)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = sheetId, Name = name };

            sheets.Append(sheet);
        }

        public static Cell CreateCell(CellValues type, string value)
        {
            var cell = new Cell();

            cell.DataType = type;
            cell.CellValue = new CellValue(value);

            return cell;
        }

        public static string GetCellValue(Row row, int index, SharedStringTable sharedString)
        {
            if (row != null && row.ChildElements.Count > index)
            {
                var cell = (Cell)row.ChildElements[index];
                if (cell.DataType != null && sharedString != null
                && cell.DataType.HasValue && cell.DataType == CellValues.SharedString
                && int.Parse(cell.CellValue.InnerText) < sharedString.ChildElements.Count)
                {
                    return sharedString.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText;
                }
                else
                {
                    return cell.CellValue.InnerText;
                }
            }

            return null;
        }
    }
}
