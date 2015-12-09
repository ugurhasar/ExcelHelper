using DocumentFormat.OpenXml.Spreadsheet;

namespace Quick.Assistant.Excel
{
    public class CellInfo
    {
        public CellInfo()
        {
        }

        public CellInfo(uint rowIndex, string columnName, CellValues dataType, string value)
        {
            this.RowIndex = rowIndex;
            this.ColumnName = columnName;
            this.DataType = dataType;
            this.Value = value;
        }

        public CellInfo(uint rowIndex, string columnName, string value)
        {
            this.RowIndex = rowIndex;
            this.ColumnName = columnName;
            this.DataType = CellValues.InlineString;
            this.Value = value;
        }

        public uint RowIndex { get; set; }
        public string ColumnName { get; set; }
        public CellValues DataType { get; set; }
        public string Value { get; set; }
    }
}
