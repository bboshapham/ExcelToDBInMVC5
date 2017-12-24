using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelToDbase.WEB.Models
{
    public class ModelReport
    {
        public int ReportID { get; set; }

        public int RegionID { get; set; }

        public string RegionName { get; set; }

        public int CompanyID { get; set; }

        public string CompanyName { get; set; }

        public int OilID { get; set; }

        public string OilName { get; set; }

        public decimal Value { get; set; }

        public DateTime CDate { get; set; }

    }
}