using System;
using System.IO;
using System.Runtime.Serialization;
using FlexCel.XlsAdapter;
using FlexCelReport;

namespace FlexCelReport
{
    [Serializable, DataContract]
    public abstract partial class ExcelReportT<T> : ExcelReportBase
    {
        public sealed override Type FilterType => typeof(T);

        #region public Object Filter


        public override Stream OnLoadTemplate(object filter, bool throwIfError)
        {
            return this.OnLoadTemplate((T)filter, throwIfError);
        }

        public override void OnBuildTemplate(ExcelReportBase.XlsReportResult xlsResult, bool throwIfError, object filter)
        {
            this.OnBuildTemplate(xlsResult, throwIfError, (T)filter);
        }

        public sealed override bool OnBuild(XlsReportResult xlsResult, bool throwIfError, object filter, bool postExecution)
        {
            return postExecution ?
                this.OnPostBuild(xlsResult, throwIfError, (T)filter) :
                this.OnBuild(xlsResult, throwIfError, (T)filter);
        }

        public sealed override bool OnLoad(FlexCel.Report.FlexCelReport flexcelReport, object filter, bool postExecution)
        {
            return postExecution ?
                this.OnPostLoad(flexcelReport, (T)filter) :
                this.OnLoad(flexcelReport, (T)filter);
        }

        public sealed override object OnFinalize(object reportDocument, bool throwIfError, object filter)
        {
            return this.OnFinalize((XlsReportResult)reportDocument, throwIfError, (T)filter);
        }

        public sealed override void OnAsposeFinalize(Aspose.Cells.Workbook workbook, object filter)
        {
            this.OnAsposeFinalize(workbook, (T)filter);
        }

        #endregion

        #region public Typed Filter (Helped for overiables)

        /// <summary>
        /// Load ra template stream, trả về null nếu mẫu/report được build tự động
        /// </summary>
        protected virtual Stream OnLoadTemplate(T filter, bool throwIfError)
        {
            return base.OnLoadTemplate(filter, throwIfError);
        }

        /// <summary>
        /// Hàm này gọi để build mẫu tự động nếu mẫu chuẩn ko có - hàm OnLoadTemplate trả về NULL
        /// </summary>
        protected virtual void OnBuildTemplate(XlsReportResult xlsResult, bool throwIfError, T filter)
        {
        }

        /// <summary>
        /// Ghi đè ở lớp dưới (lưu ý không cần gọi lại lớp gốc base.OnBuild(xlsResult, modelcontext, filter))
        /// Để thực hiện xây dựng một file excel động không dùng engine báo cáo của FlexCel (chỉ dùng Excel API)
        /// Từ mẫu gốc Excel là file excel chuẩn bình thường, không sử dụng các hàm Flexcel.
        /// Mặc định hàm OnBuild sẽ được hệ thống gọi, trong hàm OnBuild này sẽ gọi tới OnLoad cho các báo cáo
        /// thông thường. Nếu ghi đè ở lớp dưới thì hàm OnLoad sẽ được bỏ qua, nên ko dùng ghi đè ở OnLoad
        /// <returns>trả về false xác định lỗi quá trình xây dựng file excel</returns>
        /// </summary>
        protected virtual bool OnBuild(XlsFile xlsResult, bool throwIfError, T filter)
        {
            return base.OnBuild((XlsReportResult)xlsResult, throwIfError, filter, false);
        }

        /// <summary>
        /// Ghi đè ở lớp dưới để thực hiện xây dựng một báo cáo flexcel chuẩn
        /// Từ mẫu gốc Excel theo mô hình mẫu báo cáo Flexcel chuẩn
        /// <remarks>
        /// flexcelReport.AddTable(((IQueryable[EntityType])query).ToFlexCelTable("TABLE_NAME"));
        /// flexcelReport.SetValue("VALUE_NAME", value);
        /// flexcelReport.SetExpression("EXPRESSION_NAME", expression);
        /// flexcelReport.SetUserFunction("FUNCTION_NAME", Activator.CreateInstance(SystemSettings.LoadType({value})) as TFlexCelUserFunction);
        /// </remarks>
        /// </summary>
        /// <returns>false xác định report không được load do thiếu điều kiện tham số</returns>
        protected virtual bool OnLoad(FlexCel.Report.FlexCelReport flexcelReport, T filter)
        {
            return false;
        }

        /// <summary>
        /// Ghi đè ở lớp dưới (lưu ý không cần gọi lại lớp gốc base.OnBuild(xlsResult, modelcontext, filter))
        /// Để thực hiện xây dựng một file excel động không dùng engine báo cáo của FlexCel (chỉ dùng Excel API)
        /// Từ mẫu gốc Excel là file excel chuẩn bình thường, không sử dụng các hàm Flexcel.
        /// Mặc định hàm OnBuild sẽ được hệ thống gọi, trong hàm OnBuild này sẽ gọi tới OnLoad cho các báo cáo
        /// thông thường. Nếu ghi đè ở lớp dưới thì hàm OnLoad sẽ được bỏ qua, nên ko dùng ghi đè ở OnLoad
        /// <returns>trả về false xác định lỗi quá trình xây dựng file excel</returns>
        /// </summary>
        protected virtual bool OnPostBuild(XlsFile xlsResult, bool throwIfError, T filter)
        {
            return base.OnBuild((XlsReportResult)xlsResult, throwIfError, filter, true);
        }

        /// <summary>
        /// Ghi đè ở lớp dưới để thực hiện xây dựng một báo cáo flexcel chuẩn
        /// Từ mẫu gốc Excel theo mô hình mẫu báo cáo Flexcel chuẩn
        /// <remarks>
        /// flexcelReport.AddTable(((IQueryable[EntityType])query).ToFlexCelTable("TABLE_NAME"));
        /// flexcelReport.SetValue("VALUE_NAME", value);
        /// flexcelReport.SetExpression("EXPRESSION_NAME", expression);
        /// flexcelReport.SetUserFunction("FUNCTION_NAME", Activator.CreateInstance(SystemSettings.LoadType({value})) as TFlexCelUserFunction);
        /// </remarks>
        /// </summary>
        /// <returns>false xác định report không được load do thiếu điều kiện tham số (mặc định là không cần ghi đè)</returns>
        protected virtual bool OnPostLoad(FlexCel.Report.FlexCelReport flexcelReport, T filter)
        {
            return true;
        }

        /// <summary>
        /// Ghi đè ở lớp dưới để xử lý kết quả report dưới dạng Flexcel
        /// </summary>
        protected virtual XlsFile OnFinalize(XlsReportResult result, bool throwIfError, T filter)
        {
            return (XlsFile)base.OnFinalize(result, throwIfError, filter);
        }

        /// <summary>
        /// Ghi đè ở lớp dưới để xử lý kết quả report dưới dạng ASPOSE excel
        /// trước đó phải ghi đè AllowAsposeFinalize trả về true để cho phép xử lý ASPOSE
        /// </summary>
        protected virtual void OnAsposeFinalize(Aspose.Cells.Workbook workbook, T filter)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
