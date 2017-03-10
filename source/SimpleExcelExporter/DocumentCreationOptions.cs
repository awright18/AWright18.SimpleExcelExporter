using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace SimpleExcelExporter
{
    public class ExcelDocumentCreationOptions
    {
        public string WorksheetName { get; set; } = "Sheet1";

        public bool IncludeHeaderRow { get; set; } = true;

        private List<string> _ignoredColumns;

        public IEnumerable<string> IgnoredColumns
        {
            get { return _ignoredColumns; }
        }

        private List<string> _hiddenColumns;

        public IEnumerable<string> HiddenColumns
        {
            get { return _hiddenColumns; }
        }

        private Dictionary<string,string> _columnMappings = new Dictionary<string, string>();

        public IEnumerable<KeyValuePair<string, string>> ColumnMappings
        {
            get { return _columnMappings; }
        }

        public Action<ExcelPackage> ExecuteAfterDocumentCreated { get; set; }

        public static ExcelDocumentCreationOptions Default(string fileName)
        {
            return new ExcelDocumentCreationOptions(fileName);
        }

        public ExcelDocumentCreationOptions(string fileName): this()
        {
            ExecuteAfterDocumentCreated = package => { package.SaveAs(new FileInfo(fileName)); };
        }

        public ExcelDocumentCreationOptions()
        {
            _ignoredColumns = new List<string>();
            _hiddenColumns = new List<string>();
        }

        public void HideColumn(string columnName)
        {
            if (columnName == null)
            {
                throw new ArgumentNullException("columnName");
            }
            _hiddenColumns.Add(columnName);
        }

        public void DoNotWriteColumn(string columnName)
        {
            if (columnName == null)
            {
                throw new ArgumentNullException("columnName");
            }
            _ignoredColumns.Add(columnName);
        }

        public void HideColumns(IEnumerable<string> columnNames)
        {
            if (columnNames == null)
            {
                throw new ArgumentNullException("columnNames");
            }

            _hiddenColumns.AddRange(columnNames);
        }

        public void DoNotWriteColumns(IEnumerable<string> columnNames)
        {
            if (columnNames == null)
            {
                throw new ArgumentNullException("columnNames");
            }

            _ignoredColumns.AddRange(columnNames);
        }

  

        public void RenameColumn(string originalName, string newName)
        {
            _columnMappings.Add(originalName,newName);
        }

        
    }
}
