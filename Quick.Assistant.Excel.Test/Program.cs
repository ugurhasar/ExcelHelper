using DocumentFormat.OpenXml;
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
            //WriteExcel();
            //UpdateExcel();
            ReadExcel();
        }

        private static void WriteExcel()
        {
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
                }
            };

            File.WriteAllBytes("D:\\created.xlsx", excelProcess.Write(sheets));
        }

        private static void UpdateExcel()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Stream templateFile = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Quick.Assistant.Excel.Test.Embeds.template.xlsx");
                templateFile.CopyTo(stream);

                ExcelProcess excelProcess = new ExcelProcess();

                List<SheetInfo> sheets = new List<SheetInfo>()
                {
                    new SheetInfo() {
                        Name = "sheet", 
                        Cells = new List<CellInfo>(){ 
                            new CellInfo(){
                                ColumnName = "A", DataType = CellValues.InlineString, RowIndex = 3, Value = "test1"},                        
                            new CellInfo(){
                                ColumnName = "A", DataType = CellValues.InlineString, RowIndex = 2, Value = "test2"}
                            }
                        },
                    new SheetInfo() { 
                        Name = "sheet_1", 
                        Cells = new List<CellInfo>(){ 
                            new CellInfo(){
                                ColumnName = "A", DataType = CellValues.InlineString, RowIndex = 3, Value = DateTime.Now.ToShortDateString()}
                            }
                        },
                    new SheetInfo() { 
                        Name = "Sheet3", 
                        Cells = new List<CellInfo>(){ 
                            new CellInfo(){
                                ColumnName = "A", DataType = CellValues.InlineString, RowIndex = 1, Value = "Success"}
                            }
                        }
                };


                File.WriteAllBytes("D:\\updated.xlsx", excelProcess.Update(stream, sheets));
            }
        }

        private static void ReadExcel()
        {
            Stream stream = File.Open("D:\\updated.xlsx", FileMode.Open);

            ExcelProcess excelProcess = new ExcelProcess();

            using(MemoryStream memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);

                List<SheetInfo> sheetInfos = excelProcess.Read(memoryStream);
            }
        }
    }
}
