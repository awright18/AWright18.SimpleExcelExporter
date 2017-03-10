using System;
using System.Collections.Generic;
using System.Linq;

namespace SimpleExcelExporter
{
    internal class ExcelExporter<TRecordType>
    {
        internal static Func<dynamic, Dictionary<int, string>> IndexRowValues = (record) =>
        {
            var propertyNames = ((TRecordType)record).GetType().GetProperties().Select(p => p.Name).ToList();
            
            var indexedProperties = new Dictionary<int, string>();
            
            for (var index = 0; index < propertyNames.Count; index++)
            {            
                var propertyName = propertyNames[index];
            
                indexedProperties.Add(index, propertyName);
            }
            
            return indexedProperties;
        };

        internal static Func<dynamic, int, object> GetValueFromRow = (record, i) =>
        {
            var value = ((TRecordType) record).GetType().GetProperties()[i].GetValue(record);

            return value;
        };
    }
}
