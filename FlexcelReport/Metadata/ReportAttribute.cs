using System;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

namespace Report.Metadata
{
    public interface IClonableAttribute
    {
        Attribute Clone();
        Attribute New();
        void Clone(Attribute clone);
    }

    [Serializable, DataContract, ComVisible(true)]
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    public sealed class ClonableAttributeAttribute : Attribute
    {
    }

    [Serializable, DataContract]
    [AttributeUsage(AttributeTargets.Class)]
    public class ReportAttribute : Attribute
    {
        /// <summary>
        /// Tên của tác vụ (nếu không cung cấp lấy mặc định là ResourceID = "[ModuleCode].(Type).[ObjectName].[ItemName].Title")
        /// </summary>
        [DataMember]
        public string Title;
        /// <summary>
        /// Nhãn của tác vụ (nếu không cung cấp lấy mặc định là ResourceID = "[ModuleCode].(Type).[ObjectName].[ItemName].ShortTitle")
        /// </summary>
        [DataMember]
        public string ShortTitle;
        /// <summary>
        /// Miêu tả (hướng dẫn) tác vụ hiển thị dạng tooltip (nếu không cung cấp lấy mặc định là ResourceID = "[ModuleCode].(Type).[ObjectName].[ItemName].Description")
        /// </summary>
        [DataMember]
        public string Description;
        /// <summary>
        /// Kiểu tệp báo cáo (được thiết lập tự động bởi hệ thống được xác định từ khai báo Report class)
        /// </summary>
        public string ReportExt;
        /// <summary>
        /// Tên tệp báo cáo gốc không có đuôi (được thiết lập tự động bởi hệ thống được xác định từ khai báo Report class)
        /// </summary>
        public string ReportName;
        /// <summary>
        /// Tên tệp báo cáo, nếu không cung cấp lấy mặc định từ resource tiếp sau là file local theo resourceName[Report]
        /// </summary>
        public string ReportNameExt;
        /// <summary>
        /// Cung cấp đường dẫn tới file mẫu báo cáo
        /// Nếu không cung cấp thì TemplateData sẽ được sử dụng thay thế
        /// </summary>
        public string ReportPath;
        /// <summary>
        /// Cung cấp thư mục tới file mẫu báo cáo
        /// Nếu không cung cấp thì TemplateData sẽ được sử dụng thay thế
        /// </summary>
        public string ReportLocation;
        /// <summary>
        /// Cung cấp dữ liệu binary được cache lại trong bộ nhớ
        /// chứa nội dung file mẫu báo cáo
        /// (thường ở trong resource của module chứa báo cáo)
        /// </summary>
        public byte[] ReportData;
        public string ID
        {
            get { return this.ReportID.ToString(); }
            set { this.ReportID = new Guid(value); }
        }

        public Guid ReportID = Guid.Empty;
    }

    [ClonableAttribute]
    [Serializable, DataContract]
    public sealed class ExcelReportAttribute : ReportAttribute
    {
        
    }
    [ClonableAttribute]
    [Serializable, DataContract]
    public sealed class WordReportAttribute : ReportAttribute
    {

    }
}
