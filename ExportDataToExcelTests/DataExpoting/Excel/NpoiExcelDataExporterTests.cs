using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExportDataToExcel.DataExpoting.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using ExportDataToExcelTests.Models;
using System.IO;

namespace ExportDataToExcel.DataExpoting.Excel.Tests
{
    [TestClass()]
    public class NpoiExcelDataExporterTests
    {
        [TestMethod()]
        public void ExportTest()
        {
            //Prepare data
            Person[] persons = new[]
            {
                new Person
                {
                    DateOfBirth = new DateTime(1983, 12, 26),
                    FirstName = "Jon",
                    LastName = "Snow"
                },
                new Person
                {
                    DateOfBirth = new DateTime(1986, 10, 23),
                    FirstName = "Daenerys",
                    LastName = "Targaryen"
                }
            };

            //Prepare metadata
            ModelMetadataBuilder<Person> personMetadataBuilder = new ModelMetadataBuilder<Person>();

            personMetadataBuilder.Property("Name")
                .HasProvider(x => x.FirstName + " " + x.LastName);

            personMetadataBuilder.Property(x => x.DateOfBirth)
                .HasColumnName("Birthday")
                .HasProvider(x => x.DateOfBirth?.ToString("MMMM dd"));

            //Export
            ExportSheetInfo personsSheet = new ExportSheetInfo
            {
                Header = "Persons",
                Items = persons,
                Metadata = personMetadataBuilder.Metadata
            };
            ExportSheetInfo[] sheets = new[] { personsSheet };

            NpoiExcelDataExporter exporter = new NpoiExcelDataExporter();
            using (MemoryStream stream = new MemoryStream())
            {
                exporter.ExportToStream(sheets, stream);
            }
        }
    }
}