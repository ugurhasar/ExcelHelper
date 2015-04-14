using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Quick.Assistant.Excel.Test
{
  class Program
  {
    static void Main(string[] args)
    {
      //MemoryStream stream = new MemoryStream();

      //Stream templateFile = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Quick.Assistant.Excel.Test.Embeds.template.xlsx");
      //templateFile.CopyTo(stream);

      //ExcelProcess excelProcess = new ExcelProcess();
      //WorksheetPart worksheetPart = excelProcess.CloneTemplate(stream);

      //excelProcess.Write(worksheetPart);
      ExcelProcess excelProcess = new ExcelProcess();
      List<SheetInfo> sheets = new List<SheetInfo>()
      {
        new SheetInfo() {
          Name = "Sheet1", 
          Cells = new List<CellInfo>(){ 
              new CellInfo(){
                  ColumnName = "A", DataType = CellValues.InlineString, RowIndex = 1, Value = "test1"}
                }
            },
        new SheetInfo() { 
          Name = "Sheet2", 
          Cells = new List<CellInfo>(){ 
              new CellInfo(){
                  ColumnName = "A", DataType = CellValues.InlineString, RowIndex = 1, Value = DateTime.Now.ToShortDateString()}
                }
            },
      };

      File.WriteAllBytes("D:\\Deneme.xlsx", excelProcess.Write(sheets));
    }
  }
}
