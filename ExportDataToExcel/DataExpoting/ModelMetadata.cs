using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportDataToExcel.DataExpoting
{
    /// <summary>
    /// Define metadata about Model that needs to be exported.
    /// </summary>
    public class ModelMetadata
    {

        /// <summary>
        /// Information about properties to export
        /// </summary>
        public IList<PropertyMetadata> Properties { get; } = new List<PropertyMetadata>();

        /// <summary>
        /// Enumerate properties that should be exported
        /// </summary>
        /// <returns>Sorted properties to export</returns>
        public IEnumerable<PropertyMetadata> EnumeratePropertiesToExport()
        {
            return Properties
                .Where(x => !x.IsIgnored)
                .OrderBy(x => x.Order);
        }

    }
}
