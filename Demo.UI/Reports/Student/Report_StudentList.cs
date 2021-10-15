using System;
using System.Collections.Generic;
using System.Linq;
using FlexCelReport;
using Demo.UI.Manger;
using Demo.UI.Models.Report;

using Report.Metadata;

namespace Demo.UI.Reports.Student
{
    using Demo.UI.Models.Class;
    using Demo.UI.Models.Student;
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
            var classModels = this.InitClassData();
            var studentModels = this.InitStudentData(classModels);

            var classData = classModels.Select((c, index) => new {  No = SpecialCharacter[index],
                                                                    Name = "Lớp " + c.Name,
                                                                    c.Id
                                                                });
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

        private List<ClassModel> InitClassData()
        {
            return new List<ClassModel>
            {
                new ClassModel {Id=11,Name = "5A"},
                new ClassModel {Id=12,Name = "5B"},
                new ClassModel {Id=13,Name = "5C"},
                new ClassModel {Id=14,Name = "5D"},
            };
        }

        private List<StudentModel> InitStudentData(List<ClassModel> classModels)
        {
            var rs = new List<StudentModel>();
            var rand = new Random();
            foreach (var  classModel in classModels)
            {
                for (int i = 1; i <= rand.Next(3, 10); i++)
                {
                    rs.Add(new StudentModel
                    {
                        Id = i,
                        Class_Id = classModel.Id,
                        Name = "Student " + i,
                        Math = rand.Next(0,10),
                        Literature = rand.Next(0, 10)
                    });
                }
            }
            return rs;
        }

        public string[] SpecialCharacter = { "I", "II","III","IV","V","VI","VII","VIII","IX","X","XI","XII","XIII","XIV","XV" };
    }

    public class ReportFilter_StudentList : FilterModelBase
    {
        public ReportFilter_StudentList()
            : base()
        {
        }
    }
}