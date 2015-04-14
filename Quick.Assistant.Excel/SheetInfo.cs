using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quick.Assistant.Excel
{
  public class SheetInfo
  {
    public string Name { get; set; }
    public List<CellInfo> Cells { get; set; }
  }
}
