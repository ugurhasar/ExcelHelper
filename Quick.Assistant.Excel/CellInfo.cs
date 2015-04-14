using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quick.Assistant.Excel
{
  public class CellInfo
  {
    public uint RowIndex { get; set; }
    public string ColumnName { get; set; }
    public CellValues DataType { get; set; }
    public string Value { get; set; }
  }
}
