using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report.Common
{
    public static class PrintUtils
    {
        #region P/Invoke - GlobalFree/Lock/Unlock

        [System.Runtime.InteropServices.DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GlobalFree(IntPtr handle);

        [System.Runtime.InteropServices.DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GlobalLock(IntPtr handle);

        [System.Runtime.InteropServices.DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GlobalUnlock(IntPtr handle);

        #endregion

        public static void CopyTo(this PrinterSettings printerSettings, PageSettings pageSettings)
        {
            if (!printerSettings.IsValid)
                return;
            var hdevmode = printerSettings.GetHdevmode(printerSettings.DefaultPageSettings);
            try
            {
                pageSettings.SetHdevmode(hdevmode);
            }
            finally
            {
                GlobalFree(hdevmode);
            }
        }

        public static void CopyFrom(this PrinterSettings printerSettings, PageSettings pageSettings)
        {
            if (!printerSettings.IsValid)
                return;
            var hdevmode = printerSettings.GetHdevmode(pageSettings);
            try
            {
                printerSettings.SetHdevmode(hdevmode);
            }
            finally
            {
                GlobalFree(hdevmode);
            }
        }
    }

    public sealed class PageSettingsMonitor
    {
        public PageSettingsMonitor(PageSettings pageSettings)
        {
            this.pageSettings = pageSettings;
            this.bounds = pageSettings.Bounds;
            this.color = pageSettings.Color;
            this.hardMarginX = pageSettings.HardMarginX;
            this.hardMarginY = pageSettings.HardMarginY;
            this.landscape = pageSettings.Landscape;
            this.margins = pageSettings.Margins;
            this.paperSize = pageSettings.PaperSize;
            this.paperSource = pageSettings.PaperSource;
            this.printableArea = pageSettings.PrintableArea;
            this.printerResolution = pageSettings.PrinterResolution;
        }

        PageSettings pageSettings;
        Rectangle bounds;
        bool color;
        float hardMarginX;
        float hardMarginY;
        bool landscape;
        Margins margins;
        PaperSize paperSize;
        PaperSource paperSource;
        RectangleF printableArea;
        PrinterResolution printerResolution;

        public bool IsChanged
        {
            get
            {
                return this.bounds != this.pageSettings.Bounds
                    || this.color != this.pageSettings.Color
                    || this.hardMarginX != this.pageSettings.HardMarginX
                    || this.hardMarginY != this.pageSettings.HardMarginY
                    || this.landscape != this.pageSettings.Landscape
                    //|| this.margins != this.pageSettings.Margins
                    || this.paperSize.Kind != this.pageSettings.PaperSize.Kind
                    || this.paperSize.Width != this.pageSettings.PaperSize.Width
                    || this.paperSize.Height != this.pageSettings.PaperSize.Height
                    || this.paperSource.Kind != this.pageSettings.PaperSource.Kind
                    || this.paperSource.SourceName != this.pageSettings.PaperSource.SourceName
                    || this.printableArea != this.pageSettings.PrintableArea
                    || this.printerResolution.Kind != this.pageSettings.PrinterResolution.Kind
                    || this.printerResolution.X != this.pageSettings.PrinterResolution.X
                    || this.printerResolution.Y != this.pageSettings.PrinterResolution.Y;
            }
        }
    }
}
