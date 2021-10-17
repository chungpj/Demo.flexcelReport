using System.Linq;
using FlexCelReport;
using Demo.UI.Manger;
using Demo.UI.Models.Report;

using Report.Metadata;
using Core.Services;
using Core.Services.Student;

namespace Demo.UI.Reports.Student
{
    using FlexCel.Report;

    [ExcelReport(ReportName = "Student_ListStudentGroupByClass"
        , ID = "E2FA8511-04E8-409B-BEDF-9BA834D3980A"
        , ReportExt = "xlsx"
        , Title = "Danh sách đánh giá học sinh")]
    public class Report_StudentList<T> : ExcelReport<T>
        where T : ReportFilter_StudentList
    {
        protected override void SetReportLocation(string temFile = null)
        {
            base.SetReportLocation(DataServicesLocator.FileManager.TemFolder);
        }

        protected override bool OnLoad(FlexCelReport flexcelReport, T filter)
        {
            var studentModels = (filter._currentService as IStudentService).GetAll();
            var classData = studentModels.Select(s => s.Class).Distinct()
                                                                .ToList()
                                                                .Select((c, index) => new {  
                                                                                            No = SpecialCharacter[index],
                                                                                            Name = "Lớp " + c.Name,
                                                                                            c.Id
                                                                                        }).OrderBy(c => c.Id).ToList();
            flexcelReport.AddTable("Class", classData);
            var studentData = studentModels.Select((c, index) => new {
                                                                        No = index,
                                                                        c.Name,
                                                                        c.Id,
                                                                        c.Class_Id,
                                                                        c.Math,
                                                                        c.Literature,
                                                                        c.PhoneNumber,
                                                                        c.Address,
                                                                        c.Average
                                                                    });
            flexcelReport.AddTable("Student", studentData);
            
            flexcelReport.SetValue("SchoolName", "THCS Nghi Lâm");
            flexcelReport.SetValue("CourseName", "K20");
            flexcelReport.SetValue("Year", "2021");
            flexcelReport.SetValue("Creator", "Nguyễn Văn A");
            return true;
        }

        public string[] SpecialCharacter = { "I", "II","III","IV","V","VI","VII","VIII","IX","X","XI","XII","XIII","XIV","XV" };
    }

    public class ReportFilter_StudentList : FilterModelBase
    {
        public ReportFilter_StudentList(IBaseService currentService)
            : base(currentService)
        {
        }
    }
}