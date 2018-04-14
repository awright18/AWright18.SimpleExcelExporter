using System;
using System.Collections.Generic;
using AWright18.SimpleExcelExporter;
using Xunit;

namespace SampleExcelExporter.Tests
{
    public class ExcelExportTests
    {

        public class Person
        {
            public string Name { get; }
            public int Age { get; }
            public DateTime BirthDate { get; }

            public Person(string name, int age, DateTime birthDate)
            {
                Name = name;
                Age = age;
                BirthDate = birthDate;
            }

            public static IEnumerable<Person> CreateTestRecords()
            {
                return new List<Person>() { new Person("Annie",23,DateTime.Now), new Person("Bob",45, DateTime.Now)};
            }
        }
 

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

        private IEnumerable<TestRecord> GetSampleTestRecords()
        {
            return new TestRecord[]
            {
                new TestRecord("Joe", "is cool"),
                new TestRecord("Suzy", "sells sea shells")
            };
        }

        [Fact]
        public void CanExportRecordsToExcel()
        {
            var testRecords = GetSampleTestRecords();

            ExcelDocumentCreater.SaveRecordsToExcelWorksheet("sample.xlsx", testRecords, (o) =>
            {
                o.WorksheetName = "Blah";
//                o.HideColumn("Name");
//                o.DoNotWriteColumn("Value");
//                o.RenameColumn("Value","Yo");
            });


            ExcelDocumentCreater.SaveRecordsToExcelWorksheet("sample.xlsx", testRecords, (o) =>
            {
                o.WorksheetName = "Blah";
                //                o.HideColumn("Name");
                //                o.DoNotWriteColumn("Value");
                //                o.RenameColumn("Value","Yo");
            });


            var personTestRecords = Person.CreateTestRecords();
            ExcelDocumentCreater.SaveRecordsToExcelWorksheet("sample.xlsx", personTestRecords, (o) =>
            {
                o.WorksheetName = "Blah";
                //                o.HideColumn("Name");
                //                o.DoNotWriteColumn("Value");
                //                o.RenameColumn("Value","Yo");
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
            testRecords.SaveRecordsToExcelWorksheet("sample.xlsx");
        }

    }
}
