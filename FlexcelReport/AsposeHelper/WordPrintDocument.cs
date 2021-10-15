using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace Report.AsposeHelper
{
    public sealed class WordPrintDocument : PrintDocument
    {
        private Document document;
        private bool direct;
        private int index;
        private int total;

        public WordPrintDocument(Document document, bool direct)
        {
            this.document = document;
            this.direct = direct;
        }

        #region Events

        protected override void OnBeginPrint(PrintEventArgs e)
        {
            base.OnBeginPrint(e);

            var printerSettings = base.PrinterSettings;
            switch (printerSettings.PrintRange)
            {
                case PrintRange.AllPages:
                    this.index = 0;
                    this.total = this.document.PageCount;
                    break;

                case PrintRange.SomePages:
                    this.index = printerSettings.FromPage - 1;
                    this.total = printerSettings.ToPage;
                    break;
                default:
                    throw new InvalidOperationException("Lỗi giá trị khoảng trang in");
            }

            this.OnSetupPrintPage(printerSettings);
        }

        protected override void OnPrintPage(PrintPageEventArgs e)
        {
            base.OnPrintPage(e);

            this.document.RenderToScale(this.index++, e.Graphics, -e.PageSettings.HardMarginX, -e.PageSettings.HardMarginY, 1f);
            e.HasMorePages = this.index < this.total;
        }

        #endregion

        #region PrintPage

        private Dictionary<int, PaperSource> printerRawKindToPaperSourceMap = null;
        private Dictionary<PaperKind, System.Drawing.Printing.PaperSize> printerPaperKindToPaperSizeMap = null;

        private void OnSetupPrintPage(PrinterSettings printerSettings)
        {
            //if (this.direct)
            //    this.printerRawKindToPaperSourceMap = GetPrinterRawKindToPaperSourceMap(printerSettings);
            //this.printerPaperKindToPaperSizeMap = GetPrinterPaperKindToPaperSizeMap(printerSettings);
        }

        //protected override void OnQueryPageSettings(QueryPageSettingsEventArgs e)
        //{
        //    base.OnQueryPageSettings(e);

        //    this.SetPageSettings(e.PageSettings, this.index);
        //}

        //protected override void OnEndPrint(PrintEventArgs e)
        //{
        //    base.OnEndPrint(e);

        //    if (this.direct)
        //        this.printerRawKindToPaperSourceMap = null;
        //    this.printerPaperKindToPaperSizeMap = null;
        //}

        #endregion

        #region PageSettings

        public void SetPageSettings(PageSettings pageSettings, int pageIndex)
        {
            var pageInfo = this.document.GetPageInfo(pageIndex);

            if (this.direct)
            {
                var paperTray = pageInfo.PaperTray;
                var printerRawKindToPaperSourceMap = this.printerRawKindToPaperSourceMap ?? GetPrinterRawKindToPaperSourceMap(pageSettings.PrinterSettings);
                pageSettings.PaperSource = printerRawKindToPaperSourceMap.ContainsKey(paperTray) ?
                    printerRawKindToPaperSourceMap[paperTray] :
                    base.PrinterSettings.DefaultPageSettings.PaperSource;
            }
            pageSettings.PaperSize = this.GetDotNetPaperSize(pageSettings.PrinterSettings, pageInfo);
            pageSettings.Landscape = pageInfo.Landscape;
            if (pageInfo.Landscape && (pageInfo.PaperSize == Aspose.Words.PaperSize.Custom))
            {
                var height = pageSettings.PaperSize.Height;
                pageSettings.PaperSize.Height = pageSettings.PaperSize.Width;
                pageSettings.PaperSize.Width = height;
            }
        }

        private System.Drawing.Printing.PaperSize GetDotNetPaperSize(PrinterSettings printerSettings, PageInfo pageInfo)
        {
            var kind = GetDotNetPaperKind(pageInfo.PaperSize);
            var printerPaperKindToPaperSizeMap = this.printerPaperKindToPaperSizeMap ?? GetPrinterPaperKindToPaperSizeMap(printerSettings);
            return printerPaperKindToPaperSizeMap.ContainsKey(kind) ?
                printerPaperKindToPaperSizeMap[kind] :
                new System.Drawing.Printing.PaperSize("Custom",
                    (int)Math.Round(((double)pageInfo.WidthInPoints / 72.0) * 100.0),
                    (int)Math.Round(((double)pageInfo.HeightInPoints / 72.0) * 100.0)
                );
        }

        private static PaperKind GetDotNetPaperKind(Aspose.Words.PaperSize paperSize)
        {
            switch (paperSize)
            {
                case Aspose.Words.PaperSize.A3:
                    return PaperKind.A3;
                case Aspose.Words.PaperSize.A4:
                    return PaperKind.A4;
                case Aspose.Words.PaperSize.A5:
                    return PaperKind.A5;
                case Aspose.Words.PaperSize.B4:
                    return PaperKind.B4;
                case Aspose.Words.PaperSize.B5:
                    return PaperKind.B5;
                case Aspose.Words.PaperSize.Custom:
                    return PaperKind.Custom;
                case Aspose.Words.PaperSize.EnvelopeDL:
                    return PaperKind.DLEnvelope;
                case Aspose.Words.PaperSize.Executive:
                    return PaperKind.Executive;
                case Aspose.Words.PaperSize.Folio:
                    return PaperKind.Folio;
                case Aspose.Words.PaperSize.Ledger:
                    return PaperKind.Ledger;
                case Aspose.Words.PaperSize.Legal:
                    return PaperKind.Legal;
                case Aspose.Words.PaperSize.Letter:
                    return PaperKind.Letter;
                case Aspose.Words.PaperSize.Paper10x14:
                    return PaperKind.Standard10x14;
                case Aspose.Words.PaperSize.Paper11x17:
                    return PaperKind.Standard11x17;
                case Aspose.Words.PaperSize.Quarto:
                    return PaperKind.Quarto;
                case Aspose.Words.PaperSize.Statement:
                    return PaperKind.Statement;
                case Aspose.Words.PaperSize.Tabloid:
                    return PaperKind.Tabloid;
                default:
                    return PaperKind.Custom;
            }
        }

        private static Dictionary<int, PaperSource> GetPrinterRawKindToPaperSourceMap(PrinterSettings printerSettings)
        {
            return printerSettings.PaperSources.Cast<PaperSource>().GroupBy(ps => ps.RawKind).ToDictionary(gps => gps.Key, gps => gps.First());
        }

        private static Dictionary<PaperKind, System.Drawing.Printing.PaperSize> GetPrinterPaperKindToPaperSizeMap(PrinterSettings printerSettings)
        {
            return printerSettings.PaperSizes.Cast<System.Drawing.Printing.PaperSize>()
                .GroupBy(ps => ps.Kind).ToDictionary(gps => gps.Key, gps => gps.First());
        }

        #endregion
    }
}
