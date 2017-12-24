using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelToDbase.BLL;
using ExcelToDbase.DAL.Models;
namespace ImportExcel.WEB.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            // int i = 0;
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {
            string message = UploadFile(excelfile);
            ViewBag.Message = BLL.Impementation.ReadExcelFile(message);

            return View("Index");
        }

        public ActionResult SelectFromDB()
        {
            string message = BLL.Impementation.SelectFromDB();

            ViewBag.Message = message;
            return View("Index");

        }

        public ActionResult ViewInsertedData()
        {
            List<Reports> listReport = new List<Reports>();
            listReport = BLL.Impementation.ViewInsertedData();
            return View(listReport);
        }

        public string UploadFile(HttpPostedFileBase excelfile)
        {
            string path = "";
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select excel file";
            }
            else
            {
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    path = Server.MapPath("~/Content/Reports/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);

                    ViewBag.Error = "File uploaded";
                }
                else
                {
                    ViewBag.Error = "File type is incorrect";
                }
            }
            return path;
        }


    }
}


