using System;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace Report.Common
{
    public sealed class PrintManager
    {
        #region Constructor

        public static readonly PrintManager Instance = new PrintManager();
        private readonly PrintDialog printDialog;
        private string defaultPrinter;
        
        private PrintManager()
        {
            this.printDialog = new PrintDialog();
            this.printDialog.UseEXDialog = true;
            this.printDialog.AllowSomePages = true;
            this.printDialog.Document = null;
            this.DefaultPrinter = this.printDialog.PrinterSettings.PrinterName;
        }

        #endregion

        #region Operators

        public string DefaultPrinter
        {
            get { return this.defaultPrinter; }
            set
            {
                if (this.defaultPrinter != value)
                {
                    this.printDialog.PrinterSettings.PrinterName = value;
                    this.defaultPrinter = this.printDialog.PrinterSettings.PrinterName;
                    if (String.IsNullOrEmpty(this.defaultPrinter))
                        this.defaultPrinter = null;
                    if (!this.disablePrinterChangedHandler && this.defaultPrinterChangedHandler != null)
                        this.defaultPrinterChangedHandler(this, EventArgs.Empty);
                }
            }
        }

        public event EventHandler DefaultPrinterChanged
        {
            add { this.defaultPrinterChangedHandler += value; }
            remove { this.defaultPrinterChangedHandler -= value; }
        }
        private EventHandler defaultPrinterChangedHandler;
        private bool disablePrinterChangedHandler;

        public PrintDocument Prepare(PrintDocument printDocument, string printerName = null)
        {
            if (printerName == null)
                printerName = this.defaultPrinter;
            if (printerName != null && printDocument.PrinterSettings.PrinterName != printerName)
                printDocument.PrinterSettings.PrinterName = printerName;
            return printDocument;
        }

        public bool Print(PrintDocument printDocument, bool direct, ref PageSettings pageSettings, string printerName = null)
        {
            var changed = false;
            var print = true;
            if (direct)
            {
                this.Prepare(printDocument, printerName);
            }
            else lock (this.printDialog)
                {
                    this.printDialog.Document = printDocument;
                    try
                    {
                        // PrintDialog: Setup
                        var printerSettings = this.printDialog.PrinterSettings;
                        if ((printerName ?? this.defaultPrinter) != null && printerSettings.PrinterName != (printerName ?? this.defaultPrinter))
                        {
                            printerSettings.PrinterName = printerName ?? this.defaultPrinter;
                            printerSettings.CopyFrom(pageSettings);
                        }

                        // PrintDialog: Show
                        var printerPageSettingsMonitor = new PageSettingsMonitor(printerSettings.DefaultPageSettings);
                        print = this.printDialog.ShowDialog() == DialogResult.OK;

                        // Update Changes: Printer
                        if (printerName == null && this.defaultPrinter != null && printerSettings.PrinterName != this.defaultPrinter)
                        {
                            changed = true;
                            this.disablePrinterChangedHandler = true;
                            try
                            {
                                this.DefaultPrinter = printerSettings.PrinterName;
                            }
                            finally
                            {
                                this.disablePrinterChangedHandler = false;
                            }
                        }

                        // Update Changes: PageSettings
                        if (changed || printerPageSettingsMonitor.IsChanged)
                        {
                            printerSettings.CopyTo(pageSettings);
                            changed = true;
                        }
                    }
                    finally
                    {
                        this.printDialog.Document = null;
                    }
                }

            if (print)
                printDocument.Print();

            return changed;
        }

        #endregion
    }
}
