using DocumentFormat.OpenXml.Spreadsheet;

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
