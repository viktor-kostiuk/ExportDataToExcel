using System.Collections.Generic;
using System.IO;

namespace ExportDataToExcel.DataExpoting.Excel
{
    public interface IExcelExporter
    {
        void ExportToStream(IEnumerable<ExportSheetInfo> exportSheets, Stream stream);
    }
}