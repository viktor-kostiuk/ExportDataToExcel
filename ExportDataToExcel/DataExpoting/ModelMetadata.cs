using System;
using System.Collections.Generic;
using System.Text;

namespace ExportDataToExcel.DataExpoting
{
    public class ModelMetadata
    {

        public IList<PropertyMetadata> Properties { get; } = new List<PropertyMetadata>();

    }
}
