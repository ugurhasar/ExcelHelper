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
    public void CloneTemplate(Stream stream)
    {
      throw new NotImplementedException();
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
          WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
          string workbookPartId = workbookPart.GetIdOfPart(worksheetPart);

          SheetData sheetData = new SheetData();
          worksheetPart.Worksheet = new Worksheet(sheetData);

          Sheet sheet = new Sheet()
            {
              Id = workbookPartId,
              SheetId = counter,
              Name = sheetInfo.Name
            };

          sheets.AppendChild(sheet);
          
          //cells ile row gruplaması yapılacak.
          Row row = new Row { RowIndex = 1 };

          foreach (var cellInfo in sheetInfo.Cells)
          {
            Cell cell = base.CreateCell(cellInfo);
            row.AppendChild(cell);
          }

          sheetData.AppendChild(row);
          counter++;
        }

        workbookPart.Workbook.Save();
      }

      return memoryStream.ToArray();
    }

    public void CreateSheet(List<CellInfo> cells)
    {

    }

    public void Read()
    {
      throw new NotImplementedException();
    }

    //public WorksheetPart GetFirstWorkSheetPart(WorkbookPart workbookPart)
    //{
    //  //Get the relationship id of the sheetname
    //  string relId = workbookPart.Workbook.Descendants<Sheet>().First().Id;
    //  //.Where(s => s.Name.Value.Equals(sheetName))
    //  //.First()
    //  //.Id;
    //  return (WorksheetPart)workbookPart.GetPartById(relId);
    //}
  }
}
