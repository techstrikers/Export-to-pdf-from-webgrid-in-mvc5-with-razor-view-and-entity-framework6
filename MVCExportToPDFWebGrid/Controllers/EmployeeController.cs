using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MVCDisplayWebGrid.Models;
using System.Web.Helpers;
using System.Text;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace MVCDisplayWebGrid.Controllers
{
    public class EmployeeController : Controller
    {
        public ActionResult Index()
        {
            EmployeeDBContext db = new EmployeeDBContext();
            return View(db.Employees.ToList());
        }

        public FileStreamResult ExportData()
        {
            EmployeeDBContext db = new EmployeeDBContext();
            List<Employee> employees = new List<Employee>();
            employees = db.Employees.ToList();

            WebGrid webGrid = new WebGrid(employees, canPage: true, rowsPerPage: 10);
            string gridData = webGrid.GetHtml(
                columns: webGrid.Columns(
                webGrid.Column(columnName: "Name", header: "Name"),
                webGrid.Column(columnName: "Designation", header: "Designation"),
                webGrid.Column(columnName: "Gender", header: "Gender"),
                webGrid.Column(columnName: "Salary", header: "Salary"),
                webGrid.Column(columnName: "City", header: "City"),
                webGrid.Column(columnName: "State", header: "State"),
                webGrid.Column(columnName: "Zip", header: "Zip")
            )).ToString();

            StringBuilder pdfExportData = new StringBuilder();
            pdfExportData.AppendLine("<html><head><style>.webgrid-table {font-family: Arial,Helvetica,sans-serif;font-size: 14px;font-weight: normal;width: 650px;display: table;border-collapse: collapse;border: solid px #C5C5C5;background-color: white;}</style>");
            pdfExportData.AppendLine("</head><body>" + gridData + "</body></html>");
            

            var bytes  = System.Text.Encoding.UTF8.GetBytes(pdfExportData.ToString());
            using(var stream = new MemoryStream(bytes))
            {
                var outStream = new MemoryStream();
                var doc= new iTextSharp.text.Document(PageSize.A4,50,50,50,50);
                var writer = PdfWriter.GetInstance(doc,outStream);
                writer.CloseStream = false;
                doc.Open();

                var xmlWorker = iTextSharp.tool.xml.XMLWorkerHelper.GetInstance();
                xmlWorker.ParseXHtml(writer,doc,stream,System.Text.Encoding.UTF8);
                doc.Close();
                outStream.Position = 0;
                return new FileStreamResult(outStream,"application/pdf");
            }
           
        }
	}
}