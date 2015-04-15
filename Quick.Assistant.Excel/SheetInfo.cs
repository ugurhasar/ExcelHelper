using System.Collections.Generic;

namespace Quick.Assistant.Excel
{
    public class SheetInfo
    {
        public SheetInfo()
        {
            this.Cells = new List<CellInfo>();
        }

        public string Name { get; set; }
        public List<CellInfo> Cells { get; set; }
    }
}
