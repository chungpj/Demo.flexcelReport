using Core.Services;
using Report.Interface;

namespace Demo.UI.Models.Report
{
    public class FilterModelBase : IFilter
    {
        public FilterModelBase(IBaseService currentService)
        {
            //DbEntities = entities;
            _currentService = currentService;
        }
        //public DbEntities DbEntities { get; set; }
        public IBaseService _currentService;
        public long? UserId { get; set; }
    }
}