using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExportDataToExcel.DataExpoting.Excel
{
    public static class ExcelExporterExtensions
    {

        public static void ExportToFile<T>(this IExcelExporter excelExporter, IEnumerable<T> items, string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException($"{nameof(fileName)} is requried", nameof(fileName));
            }

            using (FileStream fs = File.OpenWrite(fileName))
            {
                excelExporter.ExportToSteam(items, fs);
            }
        }

        public static void ExportToSteam<T>(this IExcelExporter excelExporter, IEnumerable<T> items, Stream stream)
        {
            ModelMetadataBuilder<T> modelMetadataBuilder = new ModelMetadataBuilder<T>();
            modelMetadataBuilder.InitFromProperties();

            ExportSheetInfo exportSheetInfo = new ExportSheetInfo
            {
                Header = typeof(T).Name,
                Items = items,
                Metadata = modelMetadataBuilder.Metadata
            };

            excelExporter.ExportToStream(new[] { exportSheetInfo }, stream);
        }

    }
}
