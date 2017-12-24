using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbase.BLL.Models
{
    public class TotalBalance
    {
        public string RegionName { get; set; }
        public string ProductName { get; set; }
        public decimal TotalSumma { get; set; }
        List<TotalBalance> totalBalanceList { get; set; }
    }
}
