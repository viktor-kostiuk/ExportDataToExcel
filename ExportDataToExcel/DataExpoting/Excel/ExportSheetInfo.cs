using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace ExportDataToExcel.DataExpoting.Excel
{
    public class ExportSheetInfo
    {

        public string Header { get; set; }
        public IEnumerable Items { get; set; }
        public ModelMetadata Metadata { get; set; }

    }
}
