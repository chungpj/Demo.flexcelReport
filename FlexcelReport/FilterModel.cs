using System;
using Report.Interface;

namespace FlexCelReport
{
    public class FilterModel : IFilter
    {
        public FilterModel()
        {
            TuNgay = DateTime.Now;
            DenNgay = DateTime.Now;
        }
        public DateTime? TuNgay { get; set; }
        public DateTime? DenNgay { get; set; }
    }
}
