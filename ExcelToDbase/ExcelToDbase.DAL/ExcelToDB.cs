using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using System.Linq;
using System.Text;
using ExcelToDbase.DAL.Models;
using System.Data.SqlClient;

namespace ImportExcel.DAL
{
    public class ExcelToDB
    {
        public static string ReadExcelFile(string path)
        {
            string sheetName = "";
            string connectionString = GetConnectionString(path);
            DataSet ds = new DataSet();
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }


            List<ModelReport> ReportList = InsertExcelDateToList(ds);
            string message = InsertReport(ReportList);

            return message;

        }

        public static List<ModelReport> InsertExcelDateToList(DataSet ds)
        {
            List<ModelReport> ReportList = new List<ModelReport>();

            List<Region> RegionList = GetRegions();
            List<Company> CompanyList = GetCompanies();
            List<Product> OilList = GetProducts();


            Region Region = new Region();
            Company Company = new Company();
            //get date from excel file
            string dateString = ds.Tables[0].Rows[4][0].ToString();
            DateTime date = DateTime.Parse(dateString.Split(' ').Last());

            for (int i = 6; i < ds.Tables[0].Rows.Count; i++)
            {
                string value = ds.Tables[0].Rows[i][0].ToString();

                for (int j = 1; j < 8; j++)
                {
                    if (RegionList.Exists(x => x.RegionName.Trim() == value) == true)
                    {
                        // Get Region
                        Region = RegionList.First(x => x.RegionName.Trim() == value);
                        break;
                    }
                    else if (CompanyList.Where(x => x.CompanyName.Trim() == value).Count() > 0)
                    {
                        // Get Company
                        Company = CompanyList.First(x => x.CompanyName.Trim() == value);

                        string OilName = ds.Tables[0].Rows[5][j].ToString();

                        if (OilList.Exists(x => x.ProductName.Trim() == OilName))
                        {
                            // Get Oil
                            var Oil = OilList.First(x => x.ProductName.Trim() == OilName);

                            //Get Balance
                            decimal Balance = Decimal.Parse(ds.Tables[0].Rows[i][j].ToString());

                            // Collect Report
                            ModelReport report = new ModelReport()
                            {
                                RegionID = Region.RegionId,
                                CompanyID = Company.CompanyId,
                                OilID = Oil.ProductId,
                                Value = Balance,
                                CDate = date
                            };
                            ReportList.Add(report);
                        }
                    }
                    else
                    {

                    }
                }
            }

            return ReportList;
        }

        public static string InsertReport(List<ModelReport> ReportList)
        {
            try
            {
                //List<ModelReport> reportList = GetReports();
                // List<ModelReport> ErrorList = new List<ModelReport>();
                // List<ModelReport> NewReportList = new List<ModelReport>();
                string connnectionString = GetConnectionStringDB();
                string procedureName = "InsertReport";
                int result = 0;
                /*  foreach (var item in ReportList)
                  {
                      if (reportList.Where(x => x.RegionID == item.RegionID).Count() > 0
                          && reportList.Where(x => x.CompanyID == item.CompanyID).Count() > 0
                          && reportList.Where(x => x.OilID == item.OilID).Count() > 0
                          && reportList.Where(x => x.Value == item.Value).Count() > 0
                          && reportList.Where(x => x.CDate == item.CDate).Count() > 0
                          )
                      {
                          ErrorList.Add(item);
                          //  break;
                      }
                      else
                          NewReportList.Add(item);
                  }*/

                if (ReportList.Count > 0)
                {
                    foreach (var item in ReportList)
                    {
                        using (SqlConnection connection = new SqlConnection(connnectionString))
                        {
                            connection.Open();
                            SqlCommand command = new SqlCommand(procedureName, connection);
                            command.CommandType = System.Data.CommandType.StoredProcedure;

                            SqlParameter RegionID = new SqlParameter()
                            {
                                ParameterName = "@RegionID",
                                Value = item.RegionID
                            };

                            SqlParameter CompanyID = new SqlParameter()
                            {
                                ParameterName = "@CompanyID",
                                Value = item.CompanyID
                            };


                            SqlParameter OilID = new SqlParameter()
                            {

                                ParameterName = "@OilID",
                                Value = item.OilID
                            };

                            SqlParameter Value = new SqlParameter()
                            {

                                ParameterName = "@Balance",
                                Value = item.Value
                            };
                            SqlParameter CDate = new SqlParameter()
                            {

                                ParameterName = "@CDate",
                                Value = item.CDate
                            };

                            command.Parameters.Add(RegionID);
                            command.Parameters.Add(CompanyID);
                            command.Parameters.Add(OilID);
                            command.Parameters.Add(Value);
                            command.Parameters.Add(CDate);
                            result = command.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string result = " Error in adding Order to database.";
                return result;
            }
            return "Data successfully added to database";
        }

        public static List<Reports> SelectFromDB()
        {
            return GetModelReportsList();
        }

        public static string ExportToExcel(List<TotalBalance> totalBalanceList)
        {
            string connectionString = GetConnectionString("C:\\Users\\Araika\\Documents\\Visual Studio 2015\\Projects\\ImportExcel\\ImportExcel.WEB\\Content\\Reports\\testForExcel.xlsx");
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                //conn.BeginTransaction();
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                //  cmd["TABLE_NAME"].ToString();
                cmd.CommandText = "CREATE TABLE [table2] (OilProduct VARCHAR, Region VARCHAR, TotalBalance DECIMAL );";
                cmd.ExecuteNonQuery();
                foreach (var item in totalBalanceList)
                {
                    cmd.CommandText = "INSERT INTO [table2](OilProduct,Region,TotalBalance) VALUES('" + item.RegionName + "','" + item.ProductName + "'," + item.TotalSumma + ");";
                    cmd.ExecuteNonQuery();
                }



                conn.Close();
            }

            return "Data from DB Successfuly exported to excel file";
        }

        public static List<Region> GetRegions()
        {
            Region region;
            string connectionString = GetConnectionStringDB();
            string getRegions = "GetRegions";
            List<Region> regionList = new List<Region>();
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(getRegions, connection);
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    var reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            region = new Region();
                            region.RegionId = (int)reader[0];
                            region.RegionName = (string)reader[1];
                            regionList.Add(region);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string error = e.Message;
            }
            return regionList;
        }

        private static List<Reports> GetModelReportsList()
        {
            Reports report;
            string connectionString = GetConnectionStringDB();
            string getRegions = "GetModelReportsList";
            List<Reports> regionList = new List<Reports>();
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(getRegions, connection);
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    var reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            report = new Reports();
                            report.ReportID = (int)reader["Id"];
                            report.RegionID = (int)reader["RegionId"];
                            report.CompanyID = (int)reader["CompanyId"];
                            report.OilID = (int)reader["ProductId"];
                            report.Region = new Region();
                            report.Region.RegionName = (string)reader["RegionName"];
                            report.Company = new Company();
                            report.Company.CompanyName = (string)reader["CompanyName"];
                            report.Product = new Product();
                            report.Product.ProductName = (string)reader["ProductName"];
                            report.Value = (decimal)reader["Balance"];
                            report.CDate = (DateTime)reader["CDate"];
                            regionList.Add(report);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string error = e.Message;
            }

            return regionList;
        }

        private static List<ModelReport> GetReports()
        {
            ModelReport region;
            string connectionString = GetConnectionStringDB();
            string getRegions = "GetReports";
            List<ModelReport> regionList = new List<ModelReport>();
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(getRegions, connection);
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    var reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            region = new ModelReport();
                            region.RegionID = (int)reader[0];
                            region.CompanyID = (int)reader[1];
                            region.OilID = (int)reader[2];
                            region.Value = (decimal)reader["Balance"];
                            region.CDate = (DateTime)reader["CDate"];
                            regionList.Add(region);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string error = e.Message;
            }
            return regionList;
        }

        public static List<Company> GetCompanies()
        {
            Company region;
            string connectionString = GetConnectionStringDB();
            string getRegions = "GetCompanies";
            List<Company> regionList = new List<Company>();
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(getRegions, connection);
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    var reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            region = new Company();
                            region.CompanyId = (int)reader[0];
                            region.CompanyName = (string)reader[1];
                            regionList.Add(region);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string error = e.Message;
            }
            return regionList;
        }

        public static List<Product> GetProducts()
        {
            Product region;
            string connectionString = GetConnectionStringDB();
            string getRegions = "GetOilProducts";
            List<Product> regionList = new List<Product>();
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(getRegions, connection);
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    var reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            region = new Product();
                            region.ProductId = (int)reader[0];
                            region.ProductName = (string)reader[1];
                            regionList.Add(region);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string error = e.Message;
            }
            return regionList;
        }

        private static string GetConnectionString(string fileName)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = fileName;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }


        private static string GetConnectionStringDB()
        {
            return ConfigurationManager.ConnectionStrings["ExcelConnection"].ConnectionString;

        }
    }
}
