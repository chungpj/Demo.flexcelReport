using System;
using System.Linq;
using FlexCel.Report;

namespace Report.Common
{
    public class TSqlDataTable<T> : TLinqDataTable<T>
    {
        readonly int _aDataCount;

        public TSqlDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData, int aDataCount)
            : base(aTableName, aCreatedBy, aData)
        {
            _aDataCount = aDataCount;
        }

        public TSqlDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData, int aDataCount, TLinqDataTable<T> aOrigTable, string[] aFilters)
            : base(aTableName, aCreatedBy, aData, aOrigTable, aFilters)
        {
            _aDataCount = aDataCount;
        }

        protected override VirtualDataTable CreateNewDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData, TLinqDataTable<T> aOrigTable, string[] aFilters)
        {
            //return new TSqlDataTable<T>(aTableName, aCreatedBy, aData, aOrigTable, aFilters);
            throw new NotSupportedException();
        }

        public override VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            return new TSqlDataTableState<T>(this, Data, _aDataCount, Fields, Filters, sort, masterDetailLinks, splitLink);
        }
    }
    public class TSqlDataTableState<T> : TLinqDataTableState<T>
    {
        public TSqlDataTableState(VirtualDataTable aTableData, IQueryable<T> aData, int aDataCount, TLinqFieldDefinitions<T> aFields,
            string[] aFilters,
            string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
            : base(aTableData, aData, aFields, aFilters, sort, masterDetailLinks, splitLink)
        {
            //this.GetType().BaseType.GetField("CachedCount", DynamicDelegates.ObjectBindingFlags).SetValue(this, aDataCount);
        }
    }
}
