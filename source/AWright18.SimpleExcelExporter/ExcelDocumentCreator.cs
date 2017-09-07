using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace AWright18.SimpleExcelExporter
{
    public static class ExcelDocumentCreater
    {
        public static void SaveRecordsToExcelWorksheet(string fileName, DataTable records,
            Action<ExcelDocumentCreationOptions> options = null)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            if (records == null)
            {
                throw new ArgumentNullException(nameof(records));
            }


            var documentCreationOptions = ExcelDocumentCreationOptions.Default(fileName);

            options?.Invoke(documentCreationOptions);

            using (var document = CreateExcelPackage(fileName, documentCreationOptions))
            {
                var exporter = new GenericExcelExporter(documentCreationOptions, DataTableExcelExporter.IndexRowValues,
                    DataTableExcelExporter.GetValueFromRow);

                exporter.AddRecordsToWorksheet(documentCreationOptions.WorksheetName, records, document);

                documentCreationOptions.ExecuteAfterDocumentCreated(document);
            }
        }

        public static void SaveRecordsToExcelWorksheet(this DataTable records, string fileName,
            Action<ExcelDocumentCreationOptions> options = null)
        {
            SaveRecordsToExcelWorksheet(fileName, records, options);
        }

        public static void SaveRecordsToExcelWorksheet<TRecordType>(string fileName, IEnumerable<TRecordType> records,
            Action<ExcelDocumentCreationOptions> options = null)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            if (records == null)
            {
                throw new ArgumentNullException(nameof(records));
            }

            var documentCreationOptions = ExcelDocumentCreationOptions.Default(fileName);

            options?.Invoke(documentCreationOptions);

            using (var document = CreateExcelPackage(fileName,documentCreationOptions))
            {
                var exporter = new GenericExcelExporter(documentCreationOptions,
                    ExcelExporter<TRecordType>.IndexRowValues, ExcelExporter<TRecordType>.GetValueFromRow);

                exporter.AddRecordsToWorksheet(documentCreationOptions.WorksheetName, records, document);

                documentCreationOptions.ExecuteAfterDocumentCreated(document);
            }
        }

        public static void SaveRecordsToExcelWorksheet<TRecordType>(this IEnumerable<TRecordType> records,
            string fileName,
            Action<ExcelDocumentCreationOptions> options = null)
        {
           SaveRecordsToExcelWorksheet(fileName, records, options);
        }


        public static ExcelPackage CreateExcelPackage(string filename,ExcelDocumentCreationOptions options)
        {
            if (filename == null)
            {
                throw new ArgumentNullException(nameof(filename));
            }

            if (options.OverwriteExistingDocument)
            {
                return new ExcelPackage();
            }

            if (File.Exists(filename) && options.OverwriteExistingDocument == false)
            {
                return new ExcelPackage(new FileInfo(filename));
            }

            return new ExcelPackage();          
        }
    }
}