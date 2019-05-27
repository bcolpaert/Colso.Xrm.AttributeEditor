using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
         
        // BC 22/01/2018: cellreference is always null -> use old method
        //public static string GetCellValue(Row row, int index, SharedStringTable sharedString)
        //{
        //    var column = Utilities.IndexToColumn(index+1);

        //    var cell = (Cell) row.ChildElements
        //            .Select(x => ((Cell)x))
        //            .Where(x => x != null && x.CellReference != null && x.CellReference.Value != null)
        //            .Where(x => x.CellReference.Value.StartsWith(column))
        //            .FirstOrDefault();

        //    if (cell == null)
        //        return null;

        //    if (cell.DataType != null && sharedString != null
        //        && cell.DataType.HasValue && cell.DataType == CellValues.SharedString
        //        && int.Parse(cell.CellValue.InnerText) < sharedString.ChildElements.Count)
        //    {
        //        return sharedString.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText;
        //    } 

        //    return cell.CellValue?.InnerText;
        //}

        public static string GetCellValue(Row row, int index, SharedStringTable sharedString)
        {
            if (row != null)
            {
                Cell cell = null;
                // Search cell
                foreach (Cell c in row.ChildElements)
                {
                    var cIndex = GetColumnIndex(c.CellReference);
                    if (cIndex == index)
                    {
                        cell = c;
                        break;
                    } else if (cIndex > index)
                    {
                        // we are past the index
                        break;
                    }
                }
                if (cell != null)
                {
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
            }

            return null;
        }
        private static int? GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return null;

            //remove digits
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);

                mulitplier = mulitplier * 26;
            }

            return columnNumber;
        }

    }
}
