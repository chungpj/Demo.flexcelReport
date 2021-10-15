using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Report.Metadata;
using System.Runtime.Serialization;

namespace FlexCelReport
{
    public interface IReport
    {
        ReportAttribute Attibute { get; }
    }

    [Serializable, DataContract(IsReference = true)]
    [KnownType(typeof(ExcelReportBase)), KnownType(typeof(WordReportBase))]
    public abstract class Report : IReport
    {
        #region Constructions

        protected Report()
        {
            this.Attribute = (ReportAttribute)this.GetType().GetCustomAttributes(typeof(ReportAttribute), true).Single();
            this.Attribute.ReportExt = this.Extension;
            var reportType = this.GetType();
            this.Name = reportType.Name;
            this.Code = $"Reports.{this.Name}";
        }

        public enum eType
        {
            Excel,
            Word,
            Pdf
        }

        public abstract eType Type { get; }
        public abstract string Extension { get; }
        public readonly string Module;

        /// <summary>
        /// Mã của report (dùng tham chiếu trong toàn hệ thống)
        /// </summary>
        public string Code;

        public string Name;
        public readonly ReportAttribute Attribute;
        public abstract Type FilterType { get; }

        #endregion

        #region Operations

        /// <summary>
        /// Kiểu file mặc định sử dụng trong export
        /// </summary>
        public abstract string ExportTypeDefault { get; }

        ReportAttribute IReport.Attibute => this.Attribute;

        /// <summary>
        /// Kiểm tra kiểu file có được hỗ trợ export hay không
        /// </summary>
        public abstract bool ExportTypeSupported(string type);

        /// <summary>
        /// Kiểm tra kiểu file có phải là kiểu export đặc biệt (file gốc)
        /// </summary>
        public abstract bool ExportTypeOriginal(string type);

        /// <summary>
        /// Lấy thông tin xem trước in - dữ liệu Image và kích cỡ trang - của reportDocument (sử dụng cho WEB - ko có máy in để preview)
        /// </summary>
        public abstract IEnumerable<System.Drawing.Bitmap> Preview(object reportDocument);

        /// <summary>
        /// Xuất report ra file
        /// </summary>
        public abstract void Export(object reportDocument, string path, string type);

        /// <summary>
        /// Xuất report ra stream
        /// </summary>
        public abstract void Export(object reportDocument, Stream stream, string type);

        /// <summary>
        /// Xuất report ra Image
        /// </summary>
        public abstract System.Drawing.Image Export(object reportDocument, float xCrop = 0F, float yCrop = 0F,
            float? zoom = null, System.Drawing.Color? backgroundColor = null);

        protected abstract System.Drawing.Printing.PrintDocument PrintDocumentCreate(object reportDocument, bool direct);

        protected abstract void PrintDocumentSettings(System.Drawing.Printing.PrintDocument printDocument);

        /// <summary>
        /// Tạo object report file dùng để export, print hoặc preview trên màn hình
        /// Trả về null coi như report không đủ điều kiện để hiển thị (thiếu tham số)
        /// </summary>
        public virtual System.Drawing.Printing.PrintDocument Print(object reportDocument, bool direct,
            ref System.Drawing.Printing.PageSettings pageSettings)
        {
            var printDocument = this.PrintDocumentCreate(reportDocument, direct);

            if (pageSettings != null)
            {
                printDocument.DefaultPageSettings = pageSettings;
            }
            else
            {
                this.PrintDocumentSettings(printDocument);
                pageSettings = printDocument.DefaultPageSettings;
            }

            printDocument.PrinterSettings = pageSettings.PrinterSettings;

            return printDocument;
        }

        /// <summary>
        /// Load ra template stream, trả về null nếu mẫu/report được build tự động
        /// </summary>
        public virtual Stream OnLoadTemplate(object filter, bool throwIfError)
        {
            return this.InternalLoadTemplate();
        }
        protected virtual void SetReportLocation(string tempFile = null)
        {
            this.Attribute.ReportLocation = tempFile;
        }
        public Stream InternalLoadTemplate()
        {
            this.SetReportLocation();
            if (this.Attribute.ReportPath != null)
            {
                return new FileStream(this.Attribute.ReportPath, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete);
            }
            else if (this.Attribute.ReportData != null)
            {
                return new MemoryStream(this.Attribute.ReportData);
            }
            else if (this.Attribute.ReportName != null && this.Attribute.ReportExt != null)
            {
                this.Attribute.ReportNameExt = $"{this.Attribute.ReportName}{this.Attribute.ReportExt}";
                if (this.Attribute.ReportLocation != null)
                {
                    this.Attribute.ReportPath = Path.GetFullPath(this.Attribute.ReportLocation + "\\" + this.Attribute.ReportNameExt);
                }
                else
                {
                    string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase).Replace(@"file:\", "").Replace(@"file:/", "/"), ".."), Path.DirectorySeparatorChar + "Temp" + Path.DirectorySeparatorChar);
                    this.Attribute.ReportPath = DataPath + this.Attribute.ReportNameExt;
                }
                
                return new FileStream(this.Attribute.ReportPath, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Tạo object report file dùng để export, print hoặc preview trên màn hình
        /// Trả về null coi như report không đủ điều kiện để hiển thị (thiếu tham số)
        /// </summary>
        public object Build(object filter, bool throwIfError)
        {
            using (var templateStream = this.OnLoadTemplate(filter, throwIfError))
            {

                var report = this;
                var reportDocument = report.DirectBuild(templateStream, throwIfError, filter);
                return reportDocument == null
                    ? null
                    : report.OnFinalize(reportDocument, throwIfError, filter);
            }
        }

        public abstract object DirectBuild(Stream templateStream, bool throwIfError, object filter);
        public abstract object OnFinalize(object reportDocument, bool throwIfError, object filter);

        public abstract Stream ToStream(object reportDocument);
        public abstract object FromStream(Stream reportStream);
        #endregion
    }
}
