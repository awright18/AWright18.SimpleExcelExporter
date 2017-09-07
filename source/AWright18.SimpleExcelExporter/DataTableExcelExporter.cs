using System;
using System.Collections.Generic;
using System.Data;

namespace AWright18.SimpleExcelExporter
{
    internal class DataTableExcelExporter
    {
        internal static Func<dynamic, Dictionary<int, string>> IndexRowValues = (dt) =>
        {
            var tableIndex = new Dictionary<int, string>();

            var counter = 0;

            foreach (DataColumn column in dt.Columns)
            {
                tableIndex.Add(counter, column.ColumnName);
                counter++;
            }

            return tableIndex;
        };

        internal static Func<dynamic, int, object> GetValueFromRow = (dr, i) =>
        {
            DataRow row = (DataRow) dr;

            var value = row[i];

            return value;
        };
    }
}