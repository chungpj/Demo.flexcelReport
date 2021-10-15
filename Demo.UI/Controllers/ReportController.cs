using Demo.UI.Manager;
using Demo.UI.Reports.Student;
using System.Web.Mvc;

namespace Demo.UI.Controllers
{
    public class ReportController : Controller
    {
        public ActionResult StudentList(string extension = "pdf")
        {
            var report = new Report_StudentList<ReportFilter_StudentList>();
            var filter = new ReportFilter_StudentList() { };
            var fileContent = ReportManager.ExportPdfReport(report, filter, extension);
            return new FileContentResult(fileContent, ReportManager.GetContentType(extension));
        }

        public ActionResult Download(string extension = "xlsx")
        {
            var report = new Report_StudentList<ReportFilter_StudentList>();
            var filter = new ReportFilter_StudentList() { };
            var fileContent = ReportManager.ExportPdfReport(report, filter, extension);
            return new FileContentResult(fileContent, ReportManager.GetContentType(extension));
        }
    }
}