using Report.Interface;

namespace Demo.UI.Models.Report
{
    public class FilterModelBase : IFilter
    {
        public FilterModelBase()
        {
            //DbEntities = entities;
        }
        //public DbEntities DbEntities { get; set; }
        public long? UserId { get; set; }
    }
}