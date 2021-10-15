using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using Aspose.Cells;
using Report.Common;
using FlexCel.Core;
using FlexCel.Render;
using FlexCel.Report;
using FlexCel.XlsAdapter;

namespace FlexCelReport
{
    [Serializable, DataContract]
    public abstract class ExcelReportBase : Report
    {
        public const LoadFormat ASPOSE_LOADFORMAT = LoadFormat.Xlsx;
        public const SaveFormat ASPOSE_SAVEFORMAT = SaveFormat.Xlsx;
        public sealed override eType Type => eType.Excel;
        public sealed override string Extension => ".xlsx";
        public Stream TemStream = null;
        public Stream OutStream = null;

        protected ExcelReportBase()
        {
        }

        #region Streaming

        public sealed override Stream ToStream(object reportDocument)
        {
            var stream = new MemoryStream();
            ((XlsFile)reportDocument).Save(stream);
            stream.Seek(0, SeekOrigin.Begin);
            this.OutStream = stream;
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

                var reportDocument = new XlsFile();
                reportDocument.Open(localStream);
                return reportDocument;
            }
        }
        #endregion

        #region public sealed class XlsReportResult : XlsFile
        public sealed class XlsReportResult : XlsFile
        {
            public readonly TFlexCelUserFunction CellGetter;

            public XlsReportResult()
            {
                this.CellGetter = new Cell(this);
            }

            internal sealed class Cell : TFlexCelUserFunction
            {
                public static readonly string Name = typeof(Cell).Name;

                private XlsReportResult xlsFile;

                public Cell(XlsReportResult xlsFile)
                {
                    this.xlsFile = xlsFile;
                }

                public override object Evaluate(object[] parameters)
                {
                    int row = 0, col = 0;
                    var cell = FlexCelUtils.UDFGetParameterInString(Name, parameters, 0);
                    if (String.IsNullOrWhiteSpace(cell) || !this.xlsFile.ParseExcelIndex(cell, ref row, ref col))
                        throw new ArgumentException(String.Format("Hàm {0}({1}): lỗi phân tích địa chỉ", Name, cell));
                    return this.xlsFile.GetCellValue(row, col);
                }
            }
        }

        static void InitReportFunctions(FlexCel.Report.FlexCelReport report, XlsReportResult xlsResult)
        {
            // Set parameters: Cell("A56") => Lấy ra giá trị của ô A56 trong file excel kết quả
            report.SetUserFunction(XlsReportResult.Cell.Name, xlsResult.CellGetter);
        }
        #endregion

        #region Execution - Build

        /// <summary>
        /// Ghi đè ở lớp dưới để thực hiện xây dựng một báo cáo flexcel chuẩn
        /// Từ mẫu gốc Excel theo mô hình mẫu báo cáo Flexcel chuẩn
        /// </summary>
        public abstract bool OnLoad(FlexCel.Report.FlexCelReport flexcelReport, object filter, bool postExecution);

        /// <summary>
        /// Ghi đè ở lớp dưới (lưu ý không cần gọi lại lớp gốc base.OnBuild(xlsResult, modelcontext, filter))
        /// Để thực hiện xây dựng một file excel động không dùng engine báo cáo của FlexCel (chỉ dùng Excel API)
        /// Từ mẫu gốc Excel là file excel chuẩn bình thường, không sử dụng các hàm Flexcel
        /// <returns>trả về false xác định lỗi quá trình xây dựng file excel</returns>
        /// </summary>
        public virtual bool OnBuild(XlsReportResult xlsResult, bool throwIfError, object filter, bool postExecution)
        {
            using (FlexCel.Report.FlexCelReport report = new FlexCel.Report.FlexCelReport
            {
                ResetCellSelections = true,
                ErrorsInResultFile = !throwIfError
            })
            {
                // Images & Includes
                report.GetImageData += (o, e) => GetImageData(o, e);
                report.GetInclude += (o, e) =>
                {

                };

                // Load the template data to Run the report
                if (!this.OnLoad(report, filter, postExecution))
                    return false;
                // add Functions
                InitReportFunctions(report, xlsResult);
                // Run the report

                report.Run(xlsResult);

                return true;
            }
        }

        public virtual void GetImageData(object sender, FlexCel.Report.GetImageDataEventArgs e)
        {

        }

        /// <summary>
        /// Hàm này gọi để build mẫu tự động nếu mẫu chuẩn ko có
        /// </summary>
        public abstract void OnBuildTemplate(XlsReportResult xlsResult, bool throwIfError, object filter);

        public override object DirectBuild(Stream templateStream, bool throwIfError, object filter)
        {
            try
            {
                var xlsResult = new XlsReportResult();
                if (templateStream != null)
                    xlsResult.Open(templateStream);
                else
                    this.OnBuildTemplate(xlsResult, throwIfError, filter);

                if (!this.OnBuild(xlsResult, throwIfError, filter, false))
                    return null;

                return xlsResult;
            }
            catch (FlexCelCoreException ex)
            {
                switch (ex.ErrorCode)
                {
                    case FlxErr.ErrNotConnected:
                    //throw new ErrorMessageException("Không tìm thấy tệp mẫu báo cáo", ex);
                    case FlxErr.ErrAtCell:
                    //if (throwIfError && SystemSettings.PLATFORM != SystemSettings.ePlatform.Web)
                    //{
                    //    if (usercontext != null)
                    //        Trace.ModeLog(ExceptionUtils.ToLogDetail(ex, usercontext, "rpt://" + this.GetType().FullName));
                    //    else
                    //        Trace.ModeError(ex);
                    //}
                    //if (throwIfError || SystemSettings.PLATFORM == SystemSettings.ePlatform.Web)
                    //    throw;
                    //else
                    //    throw new ErrorMessageException(ex.Message, ex);
                    default:
                        //if (throwIfError && SystemSettings.PLATFORM != SystemSettings.ePlatform.Web)
                        //{
                        //    if (usercontext != null)
                        //        Trace.ModeLog(ExceptionUtils.ToLogDetail(ex, usercontext, "rpt://" + this.GetType().FullName));
                        //    else
                        //        Trace.ModeError(ex);
                        //}
                        //throw;
                        return null;
                }
            }
            catch (FlexCelXlsAdapterException ex)
            {
                switch (ex.ErrorCode)
                {
                    case XlsErr.ErrTooManyEntries:
                    //throw new ErrorMessageException("Hệ thống không hỗ trợ báo cáo có dữ liệu quá 65535 dòng, hãy lọc dữ liệu và chạy lại báo cáo", ex);
                    default:
                        //if (throwIfError && SystemSettings.PLATFORM != SystemSettings.ePlatform.Web)
                        //{
                        //    if (usercontext != null)
                        //        Trace.ModeLog(ExceptionUtils.ToLogDetail(ex, usercontext, "rpt://" + this.GetType().FullName));
                        //    else
                        //        Trace.ModeError(ex);
                        //}
                        //throw
                        return null;
                }
            }
        }

        #endregion

        #region Execution - Final

        public override object OnFinalize(object reportDocument, bool throwIfError, object filter)
        {
            if (this.AllowAsposeFinalize
#if DEBUG && !WEB
                && throwIfError
#endif
                )
                using (var input = this.ToStream(reportDocument))
                {
                    var workbook = new Workbook(input);
                    this.OnAsposeFinalize(workbook, filter);

                    using (var output = new MemoryStream(2048))
                    {
                        workbook.Save(output, ASPOSE_SAVEFORMAT);
                        output.Seek(0, SeekOrigin.Begin);

                        var xlsResult2 = new XlsFile();
                        xlsResult2.Open(output);
                        return xlsResult2;
                    }
                }
            return reportDocument;
        }

        /// <summary>
        /// Ghi đè ở lớp dưới trả về giá trị cho phép thực hiện
        /// FinalExecution dưới dạng ASPOSE Excel hay không (mặc định không cho phép)
        /// </summary>
        protected virtual bool AllowAsposeFinalize
        {
            get { return false; }
        }

        public abstract void OnAsposeFinalize(Workbook workbook, object filter);

        #endregion

        #region Print
        protected override PrintDocument PrintDocumentCreate(object reportDocument, bool direct)
        {
            return new FlexCelPrintDocument((ExcelFile)reportDocument)
            {
                AllVisibleSheets = false,
                ResetPageNumberOnEachSheet = false
            };
        }

        protected override void PrintDocumentSettings(PrintDocument printDocument)
        {
            if (!printDocument.PrinterSettings.IsValid)
                return;
            var flexCelPrintDocument = (FlexCelPrintDocument)printDocument;

            var workbookMargins = flexCelPrintDocument.Workbook.GetPrintMargins();
            flexCelPrintDocument.DefaultPageSettings.Margins = new Margins((int)(workbookMargins.Left * 100), (int)(workbookMargins.Right * 100), (int)(workbookMargins.Top * 100), (int)(workbookMargins.Bottom * 100));
            var paperSize = flexCelPrintDocument.Workbook.PrintPaperSize == TPaperSize.Undefined ? null : flexCelPrintDocument.PrinterSettings.PaperSizes.Cast<PaperSize>().SingleOrDefault(pz => pz.Kind == (PaperKind)flexCelPrintDocument.Workbook.PrintPaperSize);
            if (paperSize != null)
                flexCelPrintDocument.DefaultPageSettings.PaperSize = paperSize;
            else
                //Trace.ModeLog(Context.BuildName(Context.Current(), null).Insert(0, '/').Insert(0, this.Code).Insert(0, @"[rpt://").AppendFormat(@"] Khổ giấy {0} không được máy in hỗ trợ (dùng khổ giấy {1} hiện hành của máy in)", flexCelPrintDocument.Workbook.PrintPaperSize, flexCelPrintDocument.DefaultPageSettings.PaperSize.Kind).ToString());
                flexCelPrintDocument.DefaultPageSettings.Landscape = (flexCelPrintDocument.Workbook.PrintOptions & TPrintOptions.Orientation) == 0;
            var printerResolution = flexCelPrintDocument.PrinterSettings.PrinterResolutions.Cast<PrinterResolution>().FirstOrDefault(ps => ps.X == flexCelPrintDocument.Workbook.PrintXResolution && ps.Y == flexCelPrintDocument.Workbook.PrintYResolution);
            if (printerResolution != null)
                flexCelPrintDocument.DefaultPageSettings.PrinterResolution = printerResolution;
            //else
            //Trace.ModeLog(Context.BuildName(Context.Current(), null).Insert(0, '/').Insert(0, this.Code).Insert(0, @"[rpt://").AppendFormat(@"] Độ phân giải {0}x{1} không được máy in hỗ trợ (dùng độ phân giải {2}:{3}x{4} hiện hành của máy in)", flexCelPrintDocument.Workbook.PrintXResolution, flexCelPrintDocument.Workbook.PrintYResolution, flexCelPrintDocument.DefaultPageSettings.PrinterResolution.Kind, flexCelPrintDocument.DefaultPageSettings.PrinterResolution.X, flexCelPrintDocument.DefaultPageSettings.PrinterResolution.Y).ToString());
        }

        #endregion

        #region Preview

        public override IEnumerable<Bitmap> Preview(object reportDocument)
        {
            using (var flexCelImgExport = new FlexCelImgExport
            {
                AllVisibleSheets = false,
                PageSize = null,
                ResetPageNumberOnEachSheet = false,
                Resolution = 96F,
                Workbook = (ExcelFile)reportDocument
            })
            {
                var exportInfo = flexCelImgExport.GetFirstPageExportInfo();
                for (int i = 0, l = exportInfo.TotalPages; i < l; i++)
                    yield return flexCelImgExport.ExportImageNext(ref exportInfo);
            }
        }

        #endregion

        #region Export

        public override string ExportTypeDefault
        {
            get { return "xlsx"; }
        }

        public override bool ExportTypeSupported(string type)
        {
            return ExportFileFunctions.ContainsKey(type);
        }

        public override bool ExportTypeOriginal(string type)
        {
            switch (type)
            {
                case "xlsx":
                case "xls":
                case "htm":
                    return true;
            }
            return false;
        }

        public override Image Export(object reportDocument, float xCrop = 0F, float yCrop = 0F, float? zoom = null, Color? backgroundColor = null)
        {
            using (var flexCelImgExport = new FlexCelImgExport
            {
                AllVisibleSheets = false,
                PageSize = null,
                ResetPageNumberOnEachSheet = false,
                Resolution = 96F,
                Workbook = (ExcelFile)reportDocument
            })
            {
                var exportInfo = flexCelImgExport.GetFirstPageExportInfo();
                return flexCelImgExport.ExportImageNext(ref exportInfo, xCrop, yCrop, zoom, backgroundColor);
            }
        }

        #region Export - File

        private static readonly Dictionary<string, Action<ExcelFile, string>> ExportFileFunctions = new Dictionary<string, Action<ExcelFile, string>>
        {
            { "xlsx", ExportXLSX },
            { "xls", ExportXLS },
            { "pdf", ExportPDF },
            { "csv", ExportCSV },
            { "htm", ExportHTM },
            { "tif", ExportTIF },
            { "txt", ExportTXT }
        };

        public sealed override void Export(object reportDocument, string path, string type)
        {
            ExportFileFunctions[type]((ExcelFile)reportDocument, path);
        }

        private static void ExportCSV(ExcelFile reportDocument, string path)
        {
            if (File.Exists(path))
                File.Delete(path);
            reportDocument.Save(path, TFileFormats.Text, ',', Encoding.UTF8);
        }

        private static void ExportHTM(ExcelFile reportDocument, string path)
        {
            using (var export = new FlexCelHtmlExport(reportDocument, true))
                export.Export(path, Path.GetDirectoryName(path));
        }

        private static void ExportTIF(ExcelFile reportDocument, string path)
        {
            using (var export = new FlexCelImgExport(reportDocument, true))
                export.SaveAsImage(path, ImageExportType.Tiff, ImageColorDepth.TrueColor);
        }

        private static void ExportPDF(ExcelFile reportDocument, string path)
        {
            using (var export = new FlexCelPdfExport(reportDocument, true))
                export.Export(path);
        }

        private static void ExportTXT(ExcelFile reportDocument, string path)
        {
            if (File.Exists(path))
                File.Delete(path);
            reportDocument.Save(path, TFileFormats.Text, '\t', Encoding.UTF8);
        }

        private static void ExportXLS(ExcelFile reportDocument, string path)
        {
            if (File.Exists(path))
                File.Delete(path);
            reportDocument.Save(path, TFileFormats.Xls);
        }

        private static void ExportXLSX(ExcelFile reportDocument, string path)
        {
            if (File.Exists(path))
                File.Delete(path);
            reportDocument.Save(path, TFileFormats.Xlsx);
        }

        #endregion

        #region Export - Stream

        private static readonly Dictionary<string, Action<ExcelFile, Stream>> ExportStreamFunctions = new Dictionary<string, Action<ExcelFile, Stream>>
        {
            { "xlsx", ExportXLSX },
            { "xls", ExportXLS },
            { "pdf", ExportPDF },
            { "csv", ExportCSV },
            { "htm", ExportHTM },
            { "tif", ExportTIF },
            { "txt", ExportTXT }
        };

        public sealed override void Export(object reportDocument, Stream stream, string type)
        {
            ExportStreamFunctions[type]((ExcelFile)reportDocument, stream);
        }

        private static void ExportCSV(ExcelFile reportDocument, Stream stream)
        {
            reportDocument.Save(stream, TFileFormats.Text, ',', Encoding.UTF8);
        }

        private static void ExportHTM(ExcelFile reportDocument, Stream stream)
        {
            using (var export = new FlexCelHtmlExport(reportDocument, true))
                export.Export(new StreamWriter(stream), null, null);
        }

        private static void ExportTIF(ExcelFile reportDocument, Stream stream)
        {
            using (var export = new FlexCelImgExport(reportDocument, true))
                export.SaveAsImage(stream, ImageExportType.Tiff, ImageColorDepth.TrueColor);
        }

        private static void ExportPDF(ExcelFile reportDocument, Stream stream)
        {
            using (var export = new FlexCelPdfExport(reportDocument, true))
                export.Export(stream);
        }

        private static void ExportTXT(ExcelFile reportDocument, Stream stream)
        {
            reportDocument.Save(stream, TFileFormats.Text, '\t', Encoding.UTF8);
        }

        private static void ExportXLS(ExcelFile reportDocument, Stream stream)
        {
            reportDocument.Save(stream, TFileFormats.Xls);
        }

        private static void ExportXLSX(ExcelFile reportDocument, Stream stream)
        {
            reportDocument.Save(stream, TFileFormats.Xlsx);
        }

        #endregion

        #endregion
    }
}
