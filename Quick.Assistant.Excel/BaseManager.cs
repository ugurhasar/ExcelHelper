using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Quick.Assistant.Excel
{
    public abstract class BaseManager
    {
        public WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

            if (!string.IsNullOrWhiteSpace(sheetName))
                sheets = sheets.Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
                return null;

            string sheetId = sheets.First().Id.Value;

            return (WorksheetPart)document.WorkbookPart.GetPartById(sheetId);
        }

        public WorksheetPart CreateSheet(WorkbookPart workbookPart, uint sheetId, string sheetName)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            string workbookPartId = workbookPart.GetIdOfPart(worksheetPart);

            SheetData sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            Sheet sheet = new Sheet()
            {
                Id = workbookPartId,
                SheetId = sheetId,
                Name = sheetName
            };

            workbookPart.Workbook.Sheets.AppendChild(sheet);

            return worksheetPart;
        }

        public Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();
        }

        public void CreateOrUpdateCell(CellInfo cellInfo, Row row)
        {
            var existingCells = row.Elements<Cell>();

            Cell cell = null;

            if (existingCells.Count() > 0)
            {
                cell = existingCells.Where(c => string.Compare(c.CellReference.Value, string.Concat(cellInfo.ColumnName, cellInfo.RowIndex), true) == 0).FirstOrDefault();

                if (cell == null)
                {
                    cell = this.CreateCell(cellInfo);
                    row.AppendChild(cell);
                }
                else
                {
                    cell.RemoveAllChildren();
                    cell.DataType = cellInfo.DataType;
                }
            }
            else
            {
                cell = this.CreateCell(cellInfo);

                row.AppendChild(cell);
            }

            this.CreateCellText(cell, cellInfo.Value);
        }

        public Cell CreateCell(CellInfo cellInfo)
        {
            return new Cell { DataType = cellInfo.DataType, CellReference = string.Concat(cellInfo.ColumnName, cellInfo.RowIndex) };
        }

        public void CreateCellText(Cell cell, string cellValue)
        {
            Text text = new Text { Text = cellValue };

            InlineString inlineString = new InlineString();
            inlineString.AppendChild(text);

            cell.AppendChild(inlineString);
        }

        public CellInfo ReadCell(SharedStringTablePart sharedStringTablePart, Cell cell)
        {
            CellInfo cellInfo = new CellInfo();

            var cellValue = cell.CellValue;

            var value = (cellValue == null) ? cell.InnerText : cellValue.Text;

            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                value = sharedStringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }

            if (string.IsNullOrWhiteSpace(value))
                return null;

            if (cell.CellReference.HasValue)
            {
                cellInfo.ColumnName = this.GetCellColumnName(cell.CellReference.Value);
                cellInfo.RowIndex = this.GetCellRowIndex(cell.CellReference.Value);
            }

            if (cell.DataType != null)
            {
                cellInfo.DataType = cell.DataType.Value;
            }

            //cellInfo.DataType = cell.DataType == null ? null : cell.DataType;
            cellInfo.Value = (value ?? string.Empty).Trim();

            return cellInfo;
        }

        private string GetCellColumnName(string cellReference)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);

            return match.Value;
        }

        private uint GetCellRowIndex(string cellReference)
        {
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellReference);

            return uint.Parse(match.Value);
        }
    }
}
