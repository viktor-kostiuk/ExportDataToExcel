using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;
using ExportDataToExcelTests.Models;

namespace ExportDataToExcel.DataExpoting.Tests
{
    [TestClass]
    public class PropertyMetadataTests
    {

        [TestMethod]
        public void TestGettingLastName()
        {
            PropertyInfo lastNameProperty = typeof(Person).GetProperty("LastName");

            PropertyMetadata lastNamePropertyMetadata = new PropertyMetadata
            {
                Property = lastNameProperty,
                PropertyName = lastNameProperty.Name,
                DisplayName = "Last Name"
            };

            Person person = new Person
            {
                DateOfBirth = new DateTime(1983, 12, 26),
                FirstName = "Jon",
                LastName = "Snow"
            };

            string lastName = (string)lastNamePropertyMetadata.GetValue(person);
            Assert.AreEqual("Snow", lastName);
        }

        [TestMethod]
        public void TestGettingLastNameWithValueProvider()
        {
            PropertyMetadata lastNamePropertyMetadata = new PropertyMetadata
            {
                PropertyName = "LastName",
                DisplayName = "Last Name",
                ValueProvider = (object obj) =>
                {
                    Person p = (Person)obj;
                    return p.LastName;
                }
            };

            Person person = new Person
            {
                DateOfBirth = new DateTime(1983, 12, 26),
                FirstName = "Jon",
                LastName = "Snow"
            };

            string lastName = (string)lastNamePropertyMetadata.GetValue(person);
            Assert.AreEqual("Snow", lastName);
        }

    }
}