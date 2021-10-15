using System;
using System.Runtime.Serialization;

namespace FlexCelReport
{
    [Serializable, DataContract]
    public abstract class ExcelReport<TFilter> : ExcelReportT<TFilter>
    {
        protected ExcelReport()
        {
        }
    }
    [Serializable, DataContract]
    public abstract class ObjectExcelReport<TFilter> : ExcelReportT<TFilter>
    {
        protected ObjectExcelReport()
        {
        }
    }
    [Serializable, DataContract]
    public abstract class ObjectCollectionExcelReport<TFilter> : ExcelReportT<TFilter>
    {
        protected ObjectCollectionExcelReport()
        {
        }
    }
}
