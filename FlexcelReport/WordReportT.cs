using System;
using System.IO;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Serialization;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace FlexCelReport
{
    [Serializable, DataContract]
    public abstract class WordReportT<T> : WordReportBase
    {
        public sealed override Type FilterType => typeof(T);

        #region public Object Filter

        public sealed override bool OnBuild(Document wordResult, bool throwIfError, object filter, bool postExecution)
        {
            return postExecution ?
                this.OnPostBuild(wordResult, throwIfError, (T)filter) :
                this.OnBuild(wordResult, throwIfError, (T)filter);
        }

        public sealed override bool OnLoad(Document document, object filter, bool postExecution)
        {
            return postExecution ?
                this.OnPostLoad(document, (T)filter) :
                this.OnLoad(document, (T)filter);
        }

        public sealed override object OnFinalize(object reportDocument, bool throwIfError, object filter)
        {
            return this.OnFinalize((Document)reportDocument, throwIfError, (T)filter);
        }

        #endregion

        #region Overridable Typed Filter

        /// <summary>
        /// Ghi đè ở lớp dưới (lưu ý không cần gọi lại lớp gốc base.OnBuild(wordResult, modelcontext, filter))
        /// Để thực hiện xây dựng một file word động không dùng engine báo cáo của FlexCel (chỉ dùng Word API)
        /// Từ mẫu gốc Word là file word chuẩn bình thường, không sử dụng các hàm Word.
        /// Mặc định hàm OnBuild sẽ được hệ thống gọi, trong hàm OnBuild này sẽ gọi tới OnLoad cho các báo cáo
        /// thông thường. Nếu ghi đè ở lớp dưới thì hàm OnLoad sẽ được bỏ qua, nên ko dùng ghi đè ở OnLoad
        /// <returns>trả về false xác định lỗi quá trình xây dựng file word</returns>
        /// </summary>
        protected virtual bool OnBuild(Document wordResult, bool throwIfError, T filter)
        {
            return base.OnBuild(wordResult, throwIfError, filter, false);
        }

        /// <summary>
        /// Ghi đè ở lớp dưới để thực hiện xây dựng một báo cáo word chuẩn
        /// Từ mẫu gốc Word theo mô hình mẫu báo cáo Word chuẩn
        /// <remarks>
        /// document.AddTable(((IQueryable[EntityType])query).ToFlexCelTable("TABLE_NAME"));
        /// document.SetValue("VALUE_NAME", value);
        /// document.SetExpression("EXPRESSION_NAME", expression);
        /// document.SetUserFunction("FUNCTION_NAME", Activator.CreateInstance(SystemSettings.LoadType({value})) as TFlexCelUserFunction);
        /// </remarks>
        /// </summary>
        /// <returns>false xác định report không được load do thiếu điều kiện tham số</returns>
        protected virtual bool OnLoad(Document document, T filter)
        {
            return false;
        }

        /// <summary>
        /// Ghi đè ở lớp dưới (lưu ý không cần gọi lại lớp gốc base.OnBuild(wordResult, modelcontext, filter))
        /// Để thực hiện xây dựng một file word động không dùng engine báo cáo của FlexCel (chỉ dùng Word API)
        /// Từ mẫu gốc Word là file word chuẩn bình thường, không sử dụng các hàm Word.
        /// Mặc định hàm OnBuild sẽ được hệ thống gọi, trong hàm OnBuild này sẽ gọi tới OnLoad cho các báo cáo
        /// thông thường. Nếu ghi đè ở lớp dưới thì hàm OnLoad sẽ được bỏ qua, nên ko dùng ghi đè ở OnLoad
        /// <returns>trả về false xác định lỗi quá trình xây dựng file word</returns>
        /// </summary>
        protected virtual bool OnPostBuild(Document wordResult, bool throwIfError, T filter)
        {
            return base.OnBuild(wordResult, throwIfError, filter, true);
        }

        /// <summary>
        /// Ghi đè ở lớp dưới để thực hiện xây dựng một báo cáo word chuẩn
        /// Từ mẫu gốc Word theo mô hình mẫu báo cáo Word chuẩn
        /// <remarks>
        /// document.AddTable(((IQueryable[EntityType])query).ToFlexCelTable("TABLE_NAME"));
        /// document.SetValue("VALUE_NAME", value);
        /// document.SetExpression("EXPRESSION_NAME", expression);
        /// document.SetUserFunction("FUNCTION_NAME", Activator.CreateInstance(SystemSettings.LoadType({value})) as TFlexCelUserFunction);
        /// </remarks>
        /// </summary>
        /// <returns>false xác định report không được load do thiếu điều kiện tham số (mặc định là không cần ghi đè)</returns>
        protected virtual bool OnPostLoad(Document document, T filter)
        {
            return true;
        }

        /// <summary>
        /// Ghi đè ở lớp dưới để xử lý kết quả report dưới dạng Word
        /// </summary>
        protected virtual Document OnFinalize(Document result, bool throwIfError, T filter)
        {
            return result;
        }

        #endregion

        #region public MergeField Object Filter

        #region MergeField - Image

        /// <summary>
        /// Hàm load và tạo Image hoặc Picture rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        public sealed override Shape OnCreateImage(Document document, object filter, string mergefield, string field, bool nullable, object value)
        {
            return this.OnCreateImage(document, (T)filter, mergefield, field, nullable, value);
        }

        /// <summary>
        /// Hàm nhúng hình ảnh đã load ra dạng data stream vào documentBuilder - thay thế vào MergeField trong báo cáo
        /// Hệ thống mặc định gọi hàm InsertImage: documentBuilder.InsertImage(stream, horzPos: RelativeHorizontalPosition.Column, left: 0.0, vertPos: RelativeVerticalPosition.Paragraph, top: 0.0, width: -1.0, height: -1.0, wrapType: WrapType.Inline),
        /// báo cáo ghi đè hàm này để cung cấp thêm tham số như kích thước và layout
        /// </summary>
        public sealed override Shape OnInsertImage(DocumentBuilder documentBuilder, object filter, string mergefield, string field, object value, byte[] data)
        {
            return this.OnInsertImage(documentBuilder, (T)filter, mergefield, field, value, data);
        }

        /// <summary>
        /// Hàm load ra nội dung hình ảnh nhúng vào báo cáo dựa trên tên hình ảnh (có thể là dạng ID dưới tên - do Aspose.Words chỉ hỗ trợ dạng string)
        /// </summary>
        public sealed override byte[] OnLoadImage(Document document, object filter, string mergefield, string field, string imageName)
        {
            return this.OnLoadImage(document, (T)filter, mergefield, field, imageName);
        }

        /// <summary>
        /// Hàm lấy ra nội dung hình ảnh nhúng vào báo cáo dựa trên tên hình ảnh
        /// Độ rộng và Chiều cao ảnh có thể chỉnh sửa bằng cách sửa các đầu vào
        /// </summary>
        public sealed override byte[] OnLoadImage(Document document, object filter, string mergefield, string field, long imageID)
        {
            return this.OnLoadImage(document, (T)filter, mergefield, field, imageID);
        }

        #endregion

        #region MergeField - Html

        /// <summary>
        /// Hàm load và tạo Html rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        public sealed override void OnCreateHtml(Document document, object filter, string mergefield, string field, object value)
        {
            this.OnCreateHtml(document, (T)filter, mergefield, field, value);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Html nhúng vào báo cáo dựa trên tên của Html (hoặc chính là nội dung Html - mặc định hệ thống coi đây là nội dung của Html)
        /// </summary>
        public sealed override string OnLoadHtml(Document document, object filter, string mergefield, string field, string html)
        {
            return this.OnLoadHtml(document, (T)filter, mergefield, field, html);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Html nhúng vào báo cáo dựa trên ID của file đính kèm Html
        /// </summary>
        public sealed override string OnLoadHtml(Document document, object filter, string mergefield, string field, long htmlID)
        {
            return this.OnLoadHtml(document, (T)filter, mergefield, field, htmlID);
        }

        #endregion

        #region MergeField - Rtf

        /// <summary>
        /// Hàm load và tạo Rtf rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        public sealed override Document OnCreateRtf(Document document, object filter, string mergefield, string field, object value)
        {
            return this.OnCreateRtf(document, (T)filter, mergefield, field, value);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Rtf nhúng vào báo cáo dựa trên chuỗi đầu vào có thể là: tên (ghi đè và cung cấp bởi lớp dưới), id file đính kèm dạng chuỗi, hoặc chính nội dung của Rtf
        /// </summary>
        public sealed override Stream OnLoadRtf(Document document, object filter, string mergefield, string field, string rtf)
        {
            return this.OnLoadRtf(document, (T)filter, mergefield, field, rtf);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Rtf nhúng vào báo cáo dựa trên ID của file đính kèm Rtf
        /// </summary>
        public sealed override Stream OnLoadRtf(Document document, object filter, string mergefield, string field, long rtfID)
        {
            return this.OnLoadRtf(document, (T)filter, mergefield, field, rtfID);
        }

        #endregion

        #region MergeField - Word

        /// <summary>
        /// Hàm load và tạo Word rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        public sealed override Document OnCreateWord(Document document, object filter, string mergefield, string field, object value)
        {
            return this.OnCreateWord(document, (T)filter, mergefield, field, value);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Word nhúng vào báo cáo dựa trên chuỗi đầu vào có thể là: tên (ghi đè và cung cấp bởi lớp dưới), id file đính kèm dạng chuỗi, hoặc chính nội dung của Word
        /// </summary>
        public sealed override Stream OnLoadWord(Document document, object filter, string mergefield, string field, string word)
        {
            return this.OnLoadWord(document, (T)filter, mergefield, field, word);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Word nhúng vào báo cáo dựa trên ID của file đính kèm Word
        /// </summary>
        public sealed override Stream OnLoadWord(Document document, object filter, string mergefield, string field, long wordID)
        {
            return this.OnLoadWord(document, (T)filter, mergefield, field, wordID);
        }

        #endregion

        #endregion

        #region Overridable MergeField Typed Filter

        #region MergeField - Image

        /// <summary>
        /// Hàm load và tạo Image hoặc Picture rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        protected virtual Shape OnCreateImage(Document document, T filter, string mergefield, string field, bool nullable, object value)
        {
            return base.OnCreateImage(document, filter, mergefield, field, nullable, value);
        }

        /// <summary>
        /// Hàm nhúng hình ảnh đã load ra dạng data stream vào documentBuilder - thay thế vào MergeField trong báo cáo
        /// Hệ thống mặc định gọi hàm InsertImage: documentBuilder.InsertImage(stream, horzPos: RelativeHorizontalPosition.Column, left: 0.0, vertPos: RelativeVerticalPosition.Paragraph, top: 0.0, width: -1.0, height: -1.0, wrapType: WrapType.Inline),
        /// báo cáo ghi đè hàm này để cung cấp thêm tham số như kích thước và layout
        /// </summary>
        protected virtual Shape OnInsertImage(DocumentBuilder documentBuilder, T filter, string mergefield, string field, object value, byte[] data)
        {
            return base.OnInsertImage(documentBuilder, filter, mergefield, field, value, data);
        }

        /// <summary>
        /// Hàm load ra nội dung hình ảnh nhúng vào báo cáo dựa trên tên hình ảnh (có thể là dạng ID dưới tên - do Aspose.Words chỉ hỗ trợ dạng string)
        /// </summary>
        protected virtual byte[] OnLoadImage(Document document, T filter, string mergefield, string field, string imageName)
        {
            return base.OnLoadImage(document, filter, mergefield, field, imageName);
        }

        /// <summary>
        /// Hàm lấy ra nội dung hình ảnh nhúng vào báo cáo dựa trên tên hình ảnh
        /// Độ rộng và Chiều cao ảnh có thể chỉnh sửa bằng cách sửa các đầu vào
        /// </summary>
        protected virtual byte[] OnLoadImage(Document document, T filter, string mergefield, string field, long imageID)
        {
            return base.OnLoadImage(document, filter, mergefield, field, imageID);
        }

        #endregion

        #region MergeField - Html

        /// <summary>
        /// Hàm load và tạo Html rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        protected virtual void OnCreateHtml(Document document, T filter, string mergefield, string field, object value)
        {
            base.OnCreateHtml(document, filter, mergefield, field, value);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Html nhúng vào báo cáo dựa trên tên của Html (hoặc chính là nội dung Html - mặc định hệ thống coi đây là nội dung của Html)
        /// </summary>
        protected virtual string OnLoadHtml(Document document, T filter, string mergefield, string field, string html)
        {
            return base.OnLoadHtml(document, filter, mergefield, field, html);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Html nhúng vào báo cáo dựa trên ID của file đính kèm Html
        /// </summary>
        protected virtual string OnLoadHtml(Document document, T filter, string mergefield, string field, long htmlID)
        {
            return base.OnLoadHtml(document, filter, mergefield, field, htmlID);
        }

        #endregion

        #region MergeField - Rtf

        /// <summary>
        /// Hàm load và tạo Rtf rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        protected virtual Document OnCreateRtf(Document document, T filter, string mergefield, string field, object value)
        {
            return base.OnCreateRtf(document, filter, mergefield, field, value);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Rtf nhúng vào báo cáo dựa trên chuỗi đầu vào có thể là: tên (ghi đè và cung cấp bởi lớp dưới), id file đính kèm dạng chuỗi, hoặc chính nội dung của Rtf
        /// </summary>
        protected virtual Stream OnLoadRtf(Document document, T filter, string mergefield, string field, string rtf)
        {
            return base.OnLoadRtf(document, filter, mergefield, field, rtf);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Rtf nhúng vào báo cáo dựa trên ID của file đính kèm Rtf
        /// </summary>
        protected virtual Stream OnLoadRtf(Document document, T filter, string mergefield, string field, long rtfID)
        {
            return base.OnLoadRtf(document, filter, mergefield, field, rtfID);
        }

        #endregion

        #region MergeField - Word

        /// <summary>
        /// Hàm load và tạo Word rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        protected virtual Document OnCreateWord(Document document, T filter, string mergefield, string field, object value)
        {
            return base.OnCreateWord(document, filter, mergefield, field, value);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Word nhúng vào báo cáo dựa trên chuỗi đầu vào có thể là: tên (ghi đè và cung cấp bởi lớp dưới), id file đính kèm dạng chuỗi, hoặc chính nội dung của Word
        /// </summary>
        protected virtual Stream OnLoadWord(Document document, T filter, string mergefield, string field, string word)
        {
            return base.OnLoadWord(document, filter, mergefield, field, word);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Word nhúng vào báo cáo dựa trên ID của file đính kèm Word
        /// </summary>
        protected virtual Stream OnLoadWord(Document document, T filter, string mergefield, string field, long wordID)
        {
            return base.OnLoadWord(document, filter, mergefield, field, wordID);
        }

        #endregion

        #endregion
    }
}
