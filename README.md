# ExportDataToExcel
Provide a way to export data to Excel

Usage samples

Export to file
<code>
  Person[] persons = new[]
  {
      new Person
      {
          DateOfBirth = new DateTime(1983, 12, 26),
          FirstName = "John",
          LastName = "Smith"
      },
      new Person
      {
          DateOfBirth = new DateTime(1986, 10, 23),
          FirstName = "Hipolito",
          LastName = "Hudson"
      }
  };

  NpoiExcelDataExporter exporter = new NpoiExcelDataExporter();
  exporter.ExportToFile(persons, @"D:\Sample.xlsx");
</code>

With custom properties
<code>
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
</code>
