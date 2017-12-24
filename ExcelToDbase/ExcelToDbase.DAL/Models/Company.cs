using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbase.DAL.Models
{
   public class Company
    {
        public int CompanyId { get; set; }
        public string CompanyName { get; set; }
        public List<Company> companyList { get; set; }
    }
}
