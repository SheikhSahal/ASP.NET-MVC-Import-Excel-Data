using ASPdotNETMVCExcelImport.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;


namespace ASPdotNETMVCExcelImport.Controllers
{
    public class ProductController : Controller
    {
        // GET: Product
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase excelfile)
        {
            if(excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.error = "Please select excel file";
                return View("Index");
            }
            else
            {
                if(excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/Content/" + excelfile.FileName);
                    if(System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                        excelfile.SaveAs(path);
                    //Read Data from Excel File
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;

                    List<Product> listproduct = new List<Product>();

                    for(int row = 3; row <= range.Rows.Count; row++)
                    {
                        Product p = new Product();

                        p.id = ((Excel.Range)range.Cells[row, 1]).Text;
                        p.Name = ((Excel.Range)range.Cells[row, 2]).Text;
                        p.Price = ((Excel.Range)range.Cells[row, 3]).Text;
                        p.Quantity = ((Excel.Range)range.Cells[row, 4]).Text;

                        listproduct.Add(p);
                    }
                    ViewBag.listproduct = listproduct;

                    return View("View");
                }
                else
                {
                    ViewBag.error = "File Type is incorrect";
                    return View("Index");
                }
            }
        }
    }
}