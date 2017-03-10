using System;
using System.Collections.Generic;

namespace SimpleExcelExporter
{
    using OfficeOpenXml; // This comes from EEPLUS you can find it on nuget
    using System.Data;
    public static class ExcelDocumentCreater
    {
        public static void SaveRecordsToExcelWorksheet(string fileName, DataTable records, Action<ExcelDocumentCreationOptions> options = null)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException("fileName");
            }

            if (records == null)
            {
                throw new ArgumentNullException("records");
            }


            var documentCreationOptions = ExcelDocumentCreationOptions.Default(fileName);

            options?.Invoke(documentCreationOptions);

            using (var document = new ExcelPackage())
            {
                var exporter = new GenericExcelExporter(documentCreationOptions,DataTableExcelExporter.IndexRowValues, DataTableExcelExporter.GetValueFromRow);

                exporter.AddRecordsToWorksheet(documentCreationOptions.WorksheetName, records, document);

                documentCreationOptions.ExecuteAfterDocumentCreated(document);
            }
        }

        public static void SaveRecordsToExcelWorksheet<TRecordType>(string fileName, IEnumerable<TRecordType> records, Action<ExcelDocumentCreationOptions> options = null)
        {

            if (fileName == null)
            {
                throw new  ArgumentNullException("fileName");
            }

            if (records == null)
            {
                throw new ArgumentNullException("records");
            }

            var documentCreationOptions = ExcelDocumentCreationOptions.Default(fileName);

            options?.Invoke(documentCreationOptions);

            using (var document = new ExcelPackage())
            {
                var exporter = new GenericExcelExporter(documentCreationOptions, ExcelExporter<TRecordType>.IndexRowValues, ExcelExporter<TRecordType>.GetValueFromRow);

                exporter.AddRecordsToWorksheet(documentCreationOptions.WorksheetName, records, document);

                documentCreationOptions.ExecuteAfterDocumentCreated(document);
            }
        }
    }
}
