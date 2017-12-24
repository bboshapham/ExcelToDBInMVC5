using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbase.DAL.Models
{
    public class ModelReport
    {
        public int ReportID { get; set; }

        public int RegionID { get; set; }

        public int CompanyID { get; set; }

        public int OilID { get; set; }

        public decimal Value { get; set; }

        public DateTime CDate { get; set; }
    }
}
