using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbase.DAL.Models
{
    public class Reports
    {
        public int ReportID { get; set; }

        public int RegionID { get; set; }

        public Region Region { get; set; }

        public int CompanyID { get; set; }

        public Company Company { get; set; }

        public int OilID { get; set; }

        public Product Product { get; set; }

        public decimal Value { get; set; }

        public DateTime CDate { get; set; }

        public List<Region> regionList { get; set; }

        public List<Company> companyList { get; set; }

        public List<Product> productList { get; set; }

        public List<Reports> reportList { get; set; }

    }

    public class ModelReportsList
    {
        public List<Reports> reportsList { get; set; }
    }
}
