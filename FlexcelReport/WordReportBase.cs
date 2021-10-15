using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.Serialization;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.MailMerging;
using Report.AsposeHelper;
using Report.Common;

namespace FlexCelReport
{
    [Serializable, DataContract]
    public abstract class WordReportBase : Report
    {
        public const LoadFormat ASPOSE_LOADFORMAT = LoadFormat.Docx;
        public const SaveFormat ASPOSE_SAVEFORMAT = SaveFormat.Docx;

        public static readonly byte[] EmptyInclude = StreamHelper.Read(typeof(WordReportBase).Assembly.GetManifestResourceStream("Report.AsposeHelper.EmptyInclude.docx"));
        public static readonly byte[] BlankImage = StreamHelper.Read(typeof(WordReportBase).Assembly.GetManifestResourceStream("Report.AsposeHelper.BlankImage.png"));
        public WordReportBase()
        {
        }

        public sealed override eType Type => eType.Word;

        public sealed override string Extension => ".docx";

        #region Streaming

        public sealed override Stream ToStream(object reportDocument)
        {
            var stream = new MemoryStream();
            ((Document)reportDocument).Save(stream, ASPOSE_SAVEFORMAT);
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }

        public sealed override object FromStream(Stream reportStream)
        {
            using (reportStream)
            using (var localStream = new MemoryStream())
            {
                // Must copy to localstream for GetLength command
                // not supported from remote downloadable file
                reportStream.CopyTo(localStream);
                localStream.Seek(0, SeekOrigin.Begin);

                return new Document(localStream);
            }
        }

        #endregion

        #region Build

        /// <summary>
        /// Ghi đè ở lớp dưới để thực hiện xây dựng một báo cáo word chuẩn
        /// Từ mẫu gốc Word theo mô hình mẫu báo cáo Word chuẩn
        /// </summary>
        public abstract bool OnLoad(Document document, object filter, bool postExecution);

        /// <summary>
        /// Ghi đè ở lớp dưới (lưu ý không cần gọi lại lớp gốc base.OnBuild(wordResult, modelcontext, filter))
        /// Để thực hiện xây dựng một file word động không dùng engine báo cáo của FlexCel (chỉ dùng Word API)
        /// Từ mẫu gốc Word là file word chuẩn bình thường, không sử dụng các hàm Word
        /// <returns>trả về false xác định lỗi quá trình xây dựng file word</returns>
        /// </summary>
        public virtual bool OnBuild(Document document, bool throwIfError, object filter, bool postExecution)
        {
            // Mapping from datafields to special mergedfields
            foreach (var fieldname in document.MailMerge.GetFieldNames())
            {
                var valIndex = fieldname.IndexOf(':');
                if (valIndex > 0)
                    document.MailMerge.MappedDataFields.Add(fieldname, fieldname.Substring(valIndex + 1));
            }

            // MergeFields callback
            document.MailMerge.FieldMergingCallback = new FieldMergingCallback(this, filter);

            // Load the template data to Run the report
            if (!this.OnLoad(document, filter, postExecution))
                return false;

            return true;
        }

        public override object DirectBuild(Stream templateStream, bool throwIfError, object filter)
        {
            var document = templateStream != null ? new Document(templateStream) : new Document();

            if (!this.OnBuild(document, throwIfError, filter, false))
                return null;

            return document;
        }

        #endregion

        #region MergeField - Callback

        sealed class FieldMergingCallback : IFieldMergingCallback
        {
            private readonly WordReportBase wordReport;
            private readonly object filter;

            public FieldMergingCallback(WordReportBase wordReport, object filter)
            {
                this.wordReport = wordReport;
                this.filter = filter;
            }

            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (e.FieldValue != null && e.DocumentFieldName.Length > e.FieldName.Length)
                {
                    switch (e.DocumentFieldName.Substring(0, e.DocumentFieldName.Length - e.FieldName.Length).ToUpperInvariant())
                    {
                        //case "PIC:":
                        //    this.wordReport.OnCreateImage(e.Document, this.modelcontext, this.filter, e.DocumentFieldName, e.FieldName, false, e.FieldValue);
                        //    break;
                        //case "PIC?:":
                        //    this.wordReport.OnCreateImage(e.Document, this.modelcontext, this.filter, e.DocumentFieldName, e.FieldName, true, e.FieldValue);
                        //    break;
                        case "RICH:":
                            if (e.FieldValue is long || e.FieldValue is byte[])
                                this.wordReport.OnCreateWord(e.Document, this.filter, e.DocumentFieldName, e.FieldName, e.FieldValue);
                            else if (e.FieldValue is string && !String.IsNullOrEmpty((string)e.FieldValue))
                            {
                                if (((string)e.FieldValue).StartsWith(@"{\rtf"))
                                    this.wordReport.OnCreateRtf(e.Document, this.filter, e.DocumentFieldName, e.FieldName, e.FieldValue);
                                else
                                    this.wordReport.OnCreateHtml(e.Document, this.filter, e.DocumentFieldName, e.FieldName, e.FieldValue);
                            }
                            break;
                        case "HTML:":
                            this.wordReport.OnCreateHtml(e.Document, this.filter, e.DocumentFieldName, e.FieldName, e.FieldValue);
                            break;
                        case "RTF:":
                            this.wordReport.OnCreateRtf(e.Document, this.filter, e.DocumentFieldName, e.FieldName, e.FieldValue);
                            break;
                        case "WORD:":
                            this.wordReport.OnCreateWord(e.Document, this.filter, e.DocumentFieldName, e.FieldName, e.FieldValue);
                            break;
                    }
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                if (e.FieldValue != null)
                    this.wordReport.OnCreateImage(e.Document, this.filter, e.FieldName, e.FieldName, false, e.FieldValue);
            }
        }

        #endregion

        #region MergeField - Image

        /// <summary>
        /// Hàm load và tạo Image hoặc Picture rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        public virtual Shape OnCreateImage(Document document, object filter, string mergefield, string field, bool nullable, object value)
        {
            byte[] data;
            if (value == null)
                data = null;
            else if (value is long)
                data = this.OnLoadImage(document, filter, mergefield, field, (long)value);
            else if (value is string && !String.IsNullOrEmpty((string)value))
                data = this.OnLoadImage(document, filter, mergefield, field, (string)value);
            else if (value is byte[])
                data = (byte[])value;
            else
                data = null;
            var builder = document.ReplaceMergeField(mergefield);
            if (data == null)
                if (nullable)
                    return null;
                else
                    return this.OnInsertImage(builder, filter, mergefield, field, value, WordReportBase.BlankImage);
            try
            {
                return this.OnInsertImage(builder, filter, mergefield, field, value, data);
            }
            catch (ArgumentException)
            {
                // INPUT Image Data is wrong format
                if (nullable)
                    return null;
                return this.OnInsertImage(builder, filter, mergefield, field, value, WordReportBase.BlankImage);
            }
        }

        /// <summary>
        /// Hàm nhúng hình ảnh đã load ra dạng data stream vào documentBuilder - thay thế vào MergeField trong báo cáo
        /// Hệ thống mặc định gọi hàm InsertImage: documentBuilder.InsertImage(stream, horzPos: RelativeHorizontalPosition.Column, left: 0.0, vertPos: RelativeVerticalPosition.Paragraph, top: 0.0, width: -1.0, height: -1.0, wrapType: WrapType.Inline),
        /// báo cáo ghi đè hàm này để cung cấp thêm tham số như kích thước và layout
        /// </summary>
        public virtual Shape OnInsertImage(DocumentBuilder documentBuilder, object filter, string mergefield, string field, object value, byte[] data)
        {
            return documentBuilder.InsertImage(data, RelativeHorizontalPosition.Column, 0.0, RelativeVerticalPosition.Paragraph, 0.0, -1.0, -1.0, WrapType.Inline);
        }

        /// <summary>
        /// Hàm load ra nội dung hình ảnh nhúng vào báo cáo dựa trên tên hình ảnh (có thể là dạng ID dưới tên - do Aspose.Words chỉ hỗ trợ dạng string)
        /// </summary>
        public virtual byte[] OnLoadImage(Document document, object filter, string mergefield, string field, string imageName)
        {
            long id;
            if (!long.TryParse(imageName, out id) || id <= 0)
                return null;
            return this.OnLoadImage(document, filter, mergefield, field, id);
        }

        /// <summary>
        /// Hàm lấy ra nội dung hình ảnh nhúng vào báo cáo dựa trên tên hình ảnh
        /// Độ rộng và Chiều cao ảnh có thể chỉnh sửa bằng cách sửa các đầu vào
        /// </summary>
        public virtual byte[] OnLoadImage(Document document, object filter, string mergefield, string field, long imageID)
        {
            //return SystemFiles.LoadData(modelcontext, imageID);
            return null;
        }

        #endregion

        #region MergeField - Html

        /// <summary>
        /// Hàm load và tạo Html rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        public virtual void OnCreateHtml(Document document, object filter, string mergefield, string field, object value)
        {
            string html;
            if (value is long)
                html = this.OnLoadHtml(document, filter, mergefield, field, (long)value);
            else if (value is string && !String.IsNullOrEmpty((string)value))
                html = this.OnLoadHtml(document, filter, mergefield, field, (string)value);
            else if (value is byte[])
                html = Encoding.UTF8.GetString((byte[])value);
            else
                html = null;
            if (html != null)
                document.ReplaceMergeField(mergefield).InsertHtml(html);
        }

        /// <summary>
        /// Hàm lấy ra nội dung Html nhúng vào báo cáo dựa trên tên của Html (hoặc chính là nội dung Html - mặc định hệ thống coi đây là nội dung của Html)
        /// </summary>
        public virtual string OnLoadHtml(Document document, object filter, string mergefield, string field, string html)
        {
            return html;
        }

        /// <summary>
        /// Hàm lấy ra nội dung Html nhúng vào báo cáo dựa trên ID của file đính kèm Html
        /// </summary>
        public virtual string OnLoadHtml(Document document, object filter, string mergefield, string field, long htmlID)
        {
            return null;
        }

        #endregion

        #region MergeField - Rtf

        /// <summary>
        /// Hàm load và tạo Rtf rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        public virtual Document OnCreateRtf(Document document, object filter, string mergefield, string field, object value)
        {
            Stream stream;
            if (value is long)
                stream = this.OnLoadRtf(document, filter, mergefield, field, (long)value);
            else if (value is string && !String.IsNullOrEmpty((string)value))
                stream = this.OnLoadRtf(document, filter, mergefield, field, (string)value);
            else if (value is byte[])
                stream = new MemoryStream((byte[])value);
            else
                stream = null;

            if (stream != null)
                using (stream)
                {
                    var rtfDocument = new Document(stream, new LoadOptions { Encoding = Encoding.ASCII, LoadFormat = LoadFormat.Rtf });
                    document.ReplaceMergeField(mergefield).CurrentParagraph.InsertDocument(rtfDocument);
                    return rtfDocument;
                }

            return null;
        }

        /// <summary>
        /// Hàm lấy ra nội dung Rtf nhúng vào báo cáo dựa trên chuỗi đầu vào có thể là: tên (ghi đè và cung cấp bởi lớp dưới), id file đính kèm dạng chuỗi, hoặc chính nội dung của Rtf
        /// </summary>
        public virtual Stream OnLoadRtf(Document document, object filter, string mergefield, string field, string rtf)
        {
            long rtfID;
            if (Int64.TryParse(rtf, out rtfID))
                if (rtfID > 0)
                    return this.OnLoadRtf(document, filter, mergefield, field, rtfID);
                else
                    return null;
            return new MemoryStream(Encoding.ASCII.GetBytes(rtf));
        }

        /// <summary>
        /// Hàm lấy ra nội dung Rtf nhúng vào báo cáo dựa trên ID của file đính kèm Rtf
        /// </summary>
        public virtual Stream OnLoadRtf(Document document, object filter, string mergefield, string field, long rtfID)
        {
            return null;
        }

        #endregion

        #region MergeField - Word

        /// <summary>
        /// Hàm load và tạo Word rồi đưa vào báo cáo tại vị trí của mergefield
        /// </summary>
        public virtual Document OnCreateWord(Document document, object filter, string mergefield, string field, object value)
        {
            Stream stream;
            if (value is long)
                stream = this.OnLoadWord(document, filter, mergefield, field, (long)value);
            else if (value is string && !String.IsNullOrEmpty((string)value))
                stream = this.OnLoadWord(document, filter, mergefield, field, (string)value);
            else if (value is byte[])
                stream = new MemoryStream((byte[])value);
            else
                stream = null;

            if (stream != null)
                using (stream)
                {
                    var wordDocument = new Document(stream, new LoadOptions { LoadFormat = ASPOSE_LOADFORMAT });
                    document.ReplaceMergeField(mergefield).CurrentParagraph.InsertDocument(wordDocument);
                    return wordDocument;
                }
            return null;
        }

        /// <summary>
        /// Hàm lấy ra nội dung Word nhúng vào báo cáo dựa trên chuỗi đầu vào có thể là: tên (ghi đè và cung cấp bởi lớp dưới), id file đính kèm dạng chuỗi, hoặc chính nội dung của Word
        /// </summary>
        public virtual Stream OnLoadWord(Document document, object filter, string mergefield, string field, string word)
        {
            long wordID;
            if (Int64.TryParse(word, out wordID))
                if (wordID > 0)
                    return this.OnLoadWord(document, filter, mergefield, field, wordID);
                else
                    return null;
            return new MemoryStream(Encoding.UTF8.GetBytes(word));
        }

        /// <summary>
        /// Hàm lấy ra nội dung Word nhúng vào báo cáo dựa trên ID của file đính kèm Word
        /// </summary>
        public virtual Stream OnLoadWord(Document document, object filter, string mergefield, string field, long wordID)
        {
            //return SystemFiles.LoadStream(modelcontext, wordID);
            return null;
        }

        #endregion

        #region Print

        protected override PrintDocument PrintDocumentCreate(object reportDocument, bool direct)
        {
            return new WordPrintDocument((Document)reportDocument, direct);
        }

        protected override void PrintDocumentSettings(PrintDocument printDocument)
        {
            var wordPrintDocument = (WordPrintDocument)printDocument;
            wordPrintDocument.SetPageSettings(wordPrintDocument.DefaultPageSettings, 0);
        }

        #endregion

        #region Preview

        public override IEnumerable<Bitmap> Preview(object reportDocument)
        {
            for (int i = 0, l = ((Document)reportDocument).PageCount; i < l; i++)
            {
                var pageInfo = ((Document)reportDocument).GetPageInfo(i);
                var pageSize = new SizeF(pageInfo.WidthInPoints * 100.0f / 72.0f, pageInfo.HeightInPoints * 100.0f / 72.0f);

                using (var bitmap = new Bitmap((int)pageSize.Width, (int)pageSize.Height))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        // PageUnit is measuring in 1% inches for printing
                        graphics.PageUnit = GraphicsUnit.Display;
                        ((Document)reportDocument).RenderToSize(i, graphics, 0, 0, pageSize.Width, pageSize.Height);
                    }
                    yield return bitmap;
                }
            }
        }

        #endregion

        #region Export

        private readonly static Dictionary<string, SaveFormat> ExportFunctions = new Dictionary<string, SaveFormat>
        {
            { "docx", SaveFormat.Docx },
            { "doc", SaveFormat.Doc },
            { "rtf", SaveFormat.Rtf },
            { "pdf", SaveFormat.Pdf },
            { "htm", SaveFormat.Html },
            { "tif", SaveFormat.Tiff },
            { "txt", SaveFormat.Text },
            { "mht", SaveFormat.Mhtml }
        };

        public override string ExportTypeDefault
        {
            get { return "docx"; }
        }

        public override bool ExportTypeSupported(string type)
        {
            return ExportFunctions.ContainsKey(type);
        }

        public override bool ExportTypeOriginal(string type)
        {
            switch (type)
            {
                case "docx":
                case "doc":
                case "rtf":
                case "htm":
                case "mht":
                    return true;
            }
            return false;
        }

        public override Image Export(object reportDocument, float xCrop = 0F, float yCrop = 0F, float? zoom = null, Color? backgroundColor = null)
        {
            return ((Document)reportDocument).ExportFirstPage(xCrop, yCrop, zoom, backgroundColor);
        }

        public sealed override void Export(object reportDocument, string path, string type)
        {
            if (File.Exists(path))
                File.Delete(path);
            ((Document)reportDocument).Save(path, ExportFunctions[type]);
        }

        public override void Export(object reportDocument, Stream stream, string type)
        {
            ((Document)reportDocument).Save(stream, ExportFunctions[type]);
        }

        #endregion
    }
}

