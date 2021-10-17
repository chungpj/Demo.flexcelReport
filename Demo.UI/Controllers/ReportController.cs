using Core.Services.Student;
using Demo.UI.Manager;
using Demo.UI.Reports.Student;
using System.Web.Mvc;

namespace Demo.UI.Controllers
{
    public class ReportController : Controller
    {
        private readonly IStudentService _studentService;
        public ReportController(IStudentService studentService)
        {
            _studentService = studentService;
        }
        public ActionResult StudentList(string extension = "pdf")
        {
            var report = new Report_StudentList<ReportFilter_StudentList>();
            var filter = new ReportFilter_StudentList(_studentService) { };
            var fileContent = ReportManager.ExportPdfReport(report, filter, extension);
            return new FileContentResult(fileContent, ReportManager.GetContentType(extension));
        }

        public ActionResult Download(string extension = "xlsx")
        {
            var report = new Report_StudentList<ReportFilter_StudentList>();
            var filter = new ReportFilter_StudentList(_studentService) { };
            var fileContent = ReportManager.ExportPdfReport(report, filter, extension);
            return new FileContentResult(fileContent, ReportManager.GetContentType(extension));
        }
    }
}