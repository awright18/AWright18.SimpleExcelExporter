using System.Collections.Generic;
using Xunit;

namespace AWright18.SimpleExcelExporter.Tests
{
    public class ExcelExportTests
    {
        public class TestRecord
        {
            public TestRecord(string name, string value)
            {
                Name = name;
                Value = value;
            }

            public string Name { get; set; }
            public string Value { get; set; }
        }

        public class TestRecord2
        {
            public TestRecord2(string name, string value1, string value2)
            {
                Name = name;
                Value1 = value1;
                Value2 = value2;
            }

            public string Name { get; set; }
            public string Value1 { get; set; }
            public string Value2 { get; set; }
        }

        private IEnumerable<TestRecord> GetSampleTestRecords()
        {
            return new TestRecord[]
            {
                new TestRecord("Joe", "is cool"),
                new TestRecord("Suzy", "sells sea shells")
            };
        }

        private IEnumerable<TestRecord2> GetSampleTestRecords2()
        {
            return new TestRecord2[]
            {
                new TestRecord2("Joe", "is cool","enough"),
                new TestRecord2("Suzy", "sells sea shells","by the")
            };
        }



        [Fact]
        public void CanExportRecordsToExcel()
        {
            var testRecords = GetSampleTestRecords();

            ExcelDocumentCreater.SaveRecordsToExcelWorksheet("sample.xlsx", testRecords, (o) =>
            {
                o.WorksheetName = "Blah";
                o.HideColumn("Name");
                o.DoNotWriteColumn("Value");
                o.RenameColumn("Value","Yo");
            });
        }

        [Fact]
        public void CanExportRecordsToExcelUsingExtensionMethods()
        {
            var testRecords = GetSampleTestRecords();

            testRecords.SaveRecordsToExcelWorksheet("sample.xlsx", (o) =>
            {
                o.WorksheetName = "Blah";
                o.HideColumn("Name");
                o.DoNotWriteColumn("Value");
                o.RenameColumn("Value", "Yo");
            });
        }

        [Fact]
        public void CanExportRecordsToSameFile()
        {
            var testRecords = GetSampleTestRecords();

            testRecords.SaveRecordsToExcelWorksheet("sample.xlsx");
            testRecords.SaveRecordsToExcelWorksheet("sample.xlsx", o =>
            {
                o.OverwriteExistingDocument = false;

            });
        }

        [Fact]
        public void CanExportRecordsToSameFile2()
        {
            var testRecords = GetSampleTestRecords();
            var testRecords2 = GetSampleTestRecords2();

            testRecords.SaveRecordsToExcelWorksheet("sample.xlsx");
            testRecords2.SaveRecordsToExcelWorksheet("sample.xlsx", o =>
            {
                o.OverwriteExistingDocument = false;

            });
        }

    }
}
