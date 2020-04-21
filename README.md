# ExportDataToExcel
Provide a way to export data to Excel

<a href="https://pronetcs.blogspot.com/2020/04/export-data-to-excelin-c-hello.html">Check out article with description of exporter :-)</a>

Export to file
<pre>
<code>
  Person[] people = new[]
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
  exporter.ExportToFile(people, @"D:\Sample.xlsx");
</code>
</pre>

With custom properties
<pre>
<code>
  //Prepare data
  Person[] people = new[]
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
  ExportSheetInfo peopleSheet = new ExportSheetInfo
  {
      Header = "People",
      Items = people,
      Metadata = personMetadataBuilder.Metadata
  };
  ExportSheetInfo[] sheets = new[] { peopleSheet };

  NpoiExcelDataExporter exporter = new NpoiExcelDataExporter();
  using (MemoryStream stream = new MemoryStream())
  {
      exporter.ExportToStream(sheets, stream);
  }
</code>
</pre>
