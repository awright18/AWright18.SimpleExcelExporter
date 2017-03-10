using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace SimpleExcelExporter
{
    internal class GenericExcelExporter
    {
        private readonly Func<dynamic, Dictionary<int, string>> _indexRowValues;

        private readonly Func<dynamic, int, object> _getValueFromRow;

        private Dictionary<int, string> _indexedRowValues;

        private readonly Dictionary<int, int> _excelColumnToRowIndex;

        private readonly IEnumerable<string> _columnsToHide;

        private readonly IEnumerable<string> _columnsToIgnore;

        private Dictionary<string, string> _columnMapping;

        private bool _includeHeaderRow;

        internal GenericExcelExporter(ExcelDocumentCreationOptions options, Func<dynamic, Dictionary<int, string>> indexRowValues, Func<dynamic, int, object> getValueFromRow)
        {
            _indexRowValues = indexRowValues;

            _getValueFromRow = getValueFromRow;

            _indexedRowValues = new Dictionary<int, string>();

            _includeHeaderRow = options.IncludeHeaderRow;

            _columnsToHide = new List<string>();

            _excelColumnToRowIndex = new Dictionary<int, int>();

            if (options.HiddenColumns != null)
            {
                _columnsToHide = options.HiddenColumns;
            }

            _columnsToIgnore = new List<string>();

            if (options.IgnoredColumns != null)
            {
                _columnsToIgnore = options.IgnoredColumns;
            }

            _columnMapping = new Dictionary<string, string>();

            if (_columnMapping != null)
            {
                _columnMapping = options.ColumnMappings as Dictionary<string, string>;
            }
        }

        internal void AddRecordsToWorksheet(string worksheetName, dynamic records, ExcelPackage excelPackage)
        {
            var workSheet = excelPackage.Workbook.Worksheets[worksheetName];

            if (workSheet == null)
            {
                workSheet = excelPackage.Workbook.Worksheets.Add(worksheetName);
            }

            AddRowsToWorksheet(workSheet, records);

            AutoFitColumns(workSheet);

            HideColumns(workSheet);
        }

        private void AddRowsToWorksheet(ExcelWorksheet worksheet, dynamic records)
        {
            foreach (var record in records)
            {
                SetRowNamesIndex(record);

                SetRowToExcelIndex();

                AddHeaderRowToWorksheet(worksheet);

                AddDataRowToWorksheet(worksheet, record);
            }
        }

        private void AddDataRowToWorksheet(ExcelWorksheet worksheet, dynamic record)
        {
            var lastRowNumber = worksheet.Dimension.End.Row;

            var nextRowNumber = ++lastRowNumber;

            for (var columnNumber = 1; columnNumber <= _indexedRowValues.Count; columnNumber++)
            {
                var data = GetValue(record,columnNumber);

                worksheet.SetValue(nextRowNumber, columnNumber, data);

                if (data is DateTime)
                {
                    worksheet.Cells[nextRowNumber, columnNumber].Style.Numberformat.Format = "mm-dd-yy";
                }
            }
        }

        private void AddHeaderRowToWorksheet(ExcelWorksheet worksheet)
        {
            if (!_includeHeaderRow)
            {
                return;
            }

            var value = worksheet.GetValue(0, 0);

            if (value != null)
            {
                return;
            }

            const int headerRow = 1;

            for (int columnNumber = 1; columnNumber <= _indexedRowValues.Count; columnNumber++)
            {
                var headerName = GetHeaderName(columnNumber);

                if (_columnMapping.ContainsKey(headerName))
                {
                    headerName = _columnMapping[headerName];
                }

                worksheet.SetValue(headerRow, columnNumber, headerName);
            }
        }

        private void SetRowNamesIndex(dynamic record)
        {
            if (_indexedRowValues.Count > 0)
            {
                return;
            }

            _indexedRowValues = _indexRowValues(record);

            foreach (var column in _indexedRowValues.OrderBy(pair => pair.Key))
            {
                if (_columnsToIgnore.Contains(column.Value))
                {
                    _indexedRowValues.Remove(column.Key);
                }
            }
        }

        private void SetRowToExcelIndex()
        {
            if (_excelColumnToRowIndex.Count > 0)
            {
                return;
            }

            var counter = 1;
            foreach (var key in _indexedRowValues.Keys)
            {
                _excelColumnToRowIndex.Add(counter,key);

                counter++;
            }
        }

        private string GetHeaderName(int columnNumber)
        {
            var rowIndex = _excelColumnToRowIndex[columnNumber];

            var columnName = _indexedRowValues[rowIndex];
                
            var headerName = columnName.SeparateCamelCasingBySpaces();

            return headerName;
        }

        private object GetValue(dynamic record, int columnNumber)
        {
            var rowIndex = _excelColumnToRowIndex[columnNumber];

            var value = _getValueFromRow(record, rowIndex);

            return value;
        }

        private int GetColumNumberFromPropertyName(string propertyName)
        {
            var rowIndex =  _indexedRowValues.FirstOrDefault(p => p.Value.ToLower() == propertyName.ToLower()).Key;

            var columnNumber = _excelColumnToRowIndex.FirstOrDefault(p => p.Value == rowIndex).Key;

            return columnNumber;
        }

        private void HideColumns(ExcelWorksheet workSheet)
        {
            foreach (var property in _columnsToHide)
            {
                var columNumber = GetColumNumberFromPropertyName(property);

                ExcelColumn column = workSheet.Column(columNumber);

                column.Hidden = true;
            }
        }

        private static void AutoFitColumns(ExcelWorksheet workSheet)
        {
            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
        }

    }
}
