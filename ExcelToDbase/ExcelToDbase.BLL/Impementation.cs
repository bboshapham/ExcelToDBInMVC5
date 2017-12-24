using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Web;
using ExcelToDbase.DAL;
using ExcelToDbase.DAL.Models;
//using ExcelToDbase.BLL.Models;
namespace ImportExcel.BLL
{
    public class Impementation
    {

        public static string ReadExcelFile(string file)
        {
            return DAL.ExcelToDB.ReadExcelFile(file);
        }

        public static string SelectFromDB()
        {
            List<Reports> reportList = DAL.ExcelToDB.SelectFromDB();
            List<Region> regionList = DAL.ExcelToDB.GetRegions();
            List<Product> proList = DAL.ExcelToDB.GetProducts();
            int summaCompany = DAL.ExcelToDB.GetCompanies().Count;
            List<TotalBalance> totalList = CalculateTotalBalance(reportList, regionList, proList, summaCompany);
            string answer = DAL.ExcelToDB.ExportToExcel(totalList);
            return answer;
        }

        public static List<Reports> ViewInsertedData()
        {
            List<Reports> insertReportList = DAL.ExcelToDB.SelectFromDB();

            return insertReportList;
        }

        private static List<TotalBalance> CalculateTotalBalance(List<Reports> reportList, List<Region> regionList, List<Product> prolist, int summaCompany)
        {
            int regionId = 0;
            int productId = 0;
            decimal total = 0;
            List<TotalBalance> totalList = new List<TotalBalance>();
            TotalBalance totalB;
            foreach (var item in regionList)
            {
                regionId = item.RegionId;
                foreach (var it in prolist)
                {
                    productId = it.ProductId;
                    foreach (var samoList in reportList)
                    {
                        if (regionId == samoList.RegionID && productId == samoList.OilID)
                        {
                            total += samoList.Value;
                            //    reportList.Remove(samoList);
                        }
                        else { }
                    }
                    totalB = new TotalBalance();
                    totalB.RegionName = item.RegionName;
                    totalB.ProductName = it.ProductName;
                    totalB.TotalSumma = total;
                    totalList.Add(totalB);

                }

            }
            return totalList;

        }

    }
}
