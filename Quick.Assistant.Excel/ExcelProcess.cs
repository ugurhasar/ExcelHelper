using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quick.Assistant.Excel
{
    public class ExcelProcess : BaseManager
    {
        public byte[] Update(MemoryStream memoryStream, List<SheetInfo> sheetInfos)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(memoryStream, true))
            {
                uint counter = 1;

                foreach (var sheetInfo in sheetInfos)
                {
                    WorksheetPart worksheetPart = base.GetWorksheetPartByName(spreadsheetDocument, sheetInfo.Name);

                    if (worksheetPart == null)
                        worksheetPart = base.CreateSheet(spreadsheetDocument.WorkbookPart, counter, sheetInfo.Name);

                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.Elements<SheetData>().FirstOrDefault();

                    var rowGroups = sheetInfo.Cells.GroupBy(x => x.RowIndex);

                    foreach (var rowGroup in rowGroups)
                    {
                        Row row = base.GetRow(worksheet, rowGroup.Key);

                        if (row == null)
                        {
                            row = new Row() { RowIndex = rowGroup.Key };
                            sheetData.AppendChild(row);
                        }

                        foreach (var cellInfo in rowGroup)
                            base.CreateOrUpdateCell(cellInfo, row);
                    }

                    worksheetPart.Worksheet.Save();

                    counter++;
                }

                return memoryStream.ToArray();
            }
        }

        public List<SheetInfo> Read(MemoryStream memoryStream)
        {
            List<SheetInfo> sheetInfos = new List<SheetInfo>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(memoryStream, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;

                var sheets = workbookPart.Workbook.Descendants<Sheet>();

                foreach (var sheet in sheets)
                {
                    SheetInfo sheetInfo = new SheetInfo();

                    sheetInfo.Name = sheet.Name;

                    WorksheetPart worksheetPart = base.GetWorksheetPartByName(spreadsheetDocument, sheet.Name);

                    if (worksheetPart == null)
                        continue;

                    Worksheet worksheet = worksheetPart.Worksheet;

                    SheetData sheetData = worksheet.Elements<SheetData>().FirstOrDefault();

                    List<Row> rows = sheetData.Elements<Row>().ToList();

                    if (rows.Count == 0)
                        continue;

                    foreach (Row row in rows)
                    {
                        List<Cell> cells = row.Elements<Cell>().ToList();

                        foreach (Cell cell in cells)
                        {
                            CellInfo cellInfo = base.ReadCell(sharedStringTablePart, cell);

                            if (cellInfo == null)
                                continue;

                            sheetInfo.Cells.Add(cellInfo);
                        }
                    }

                    sheetInfos.Add(sheetInfo);
                }
            }

            return sheetInfos;
        }

        public byte[] Write(List<SheetInfo> sheetInfos)
        {
            MemoryStream memoryStream = new MemoryStream();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();

                workbookPart.Workbook = new Workbook();
                Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                uint counter = 1;

                foreach (var sheetInfo in sheetInfos)
                {
                    WorksheetPart worksheetPart = base.CreateSheet(workbookPart, counter, sheetInfo.Name);

                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.Elements<SheetData>().FirstOrDefault();

                    var rowGroups = sheetInfo.Cells.GroupBy(x => x.RowIndex);

                    foreach (var rowGroup in rowGroups)
                    {
                        Row row = new Row() { RowIndex = rowGroup.Key };
                        sheetData.AppendChild(row);

                        foreach (var cellInfo in rowGroup)
                            base.CreateOrUpdateCell(cellInfo, row);
                    }

                    counter++;
                }

                workbookPart.Workbook.Save();
            }

            return memoryStream.ToArray();
        }        
    }
}
