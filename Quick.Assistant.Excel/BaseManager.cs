using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quick.Assistant.Excel
{
  public abstract class BaseManager
  {
    public Cell CreateCell(CellInfo cellInfo)
    {
      var cell = new Cell
      {
        DataType = cellInfo.DataType,
        CellReference = string.Concat(cellInfo.ColumnName, cellInfo.RowIndex)
      };

      Text text = new Text { Text = cellInfo.Value };

      InlineString inlineString = new InlineString();
      inlineString.AppendChild(text);

      cell.AppendChild(inlineString);

      return cell;
    }
  }
}
