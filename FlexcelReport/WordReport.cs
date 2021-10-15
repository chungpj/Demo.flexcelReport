using System;
using System.Runtime.Serialization;

namespace FlexCelReport
{
    [Serializable, DataContract]
    public abstract class WordReport<TFilter> : WordReportT<TFilter>
    {
        protected WordReport()
        {
        }
    }
    [Serializable, DataContract]
    public abstract class ObjectWordReport<TFilter> : WordReportT<TFilter>
    {
        protected ObjectWordReport()
        {
        }
    }
    [Serializable, DataContract]
    public abstract class ObjectCollectionWordReport<TFilter> : WordReportT<TFilter>
    {
        protected ObjectCollectionWordReport()
        {
        }
    }
}
