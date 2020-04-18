using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExportDataToExcel.DataExpoting;
using System;
using System.Collections.Generic;
using System.Text;
using ExportDataToExcelTests.Models;
using System.Linq;

namespace ExportDataToExcel.DataExpoting.Tests
{
    [TestClass]
    public class ModelMetadataBuilderTests
    {
        [TestMethod]
        public void InitFromPropertiesTest()
        {
            ModelMetadataBuilder<Person> modelMetadataBuilder = new ModelMetadataBuilder<Person>();
            modelMetadataBuilder.InitFromProperties();

            Person person = new Person
            {
                DateOfBirth = new DateTime(1983, 12, 26),
                FirstName = "Jon",
                LastName = "Snow"
            };

            ModelMetadata metadata = modelMetadataBuilder.Metadata;
            var lastNamePropertyMetadata = metadata.Properties
                .First(x => x.PropertyName == "LastName");

            string lastName = (string)lastNamePropertyMetadata.GetValue(person);
            Assert.AreEqual("Snow", lastName);
        }

        [TestMethod]
        public void PropertyMetadataBuilderResolvingTest()
        {
            ModelMetadataBuilder<Person> modelMetadataBuilder = new ModelMetadataBuilder<Person>();
            modelMetadataBuilder.InitFromProperties();

            Person person = new Person
            {
                DateOfBirth = new DateTime(1983, 12, 26),
                FirstName = "Jon",
                LastName = "Snow"
            };

            ModelMetadata metadata = modelMetadataBuilder.Metadata;
            var lastNamePropertyMetadata = modelMetadataBuilder
                .Property(x => x.LastName)
                .PropertyMetadata;

            string lastName = (string)lastNamePropertyMetadata.GetValue(person);
            Assert.AreEqual("Snow", lastName);
        }

    }
}