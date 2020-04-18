using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ExportDataToExcel.DataExpoting
{
    /// <summary>
    /// Define inforaiton about model property that needs to be exported
    /// </summary>
    public class PropertyMetadata
    {

        /// <summary>
        /// Reference to model property.
        /// Can be null if it custom informaton to export
        /// </summary>
        public PropertyInfo Property { get; set; }

        /// <summary>
        /// Name of property. 
        /// Should be unique per ModelMetadata
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// Display name of property.
        /// This name will be used as header when export
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Function to provider value for property from Model!
        /// If create ModelMetadta for Person object, then Person will be passed as paremeter to ValueProvider
        /// </summary>
        public Func<object, object> ValueProvider { get; set; }

        /// <summary>
        /// If true, this property will not be exported.
        /// </summary>
        public bool IsIgnored { get; set; }

        /// <summary>
        /// Used to sort properties before export.
        /// </summary>
        public int? Order { get; set; }

        /// <summary>
        /// Get value of proprety
        /// </summary>
        /// <param name="obj">Model object</param>
        /// <returns>Value from model for this property</returns>
        public object GetValue(object obj) => ValueProvider != null
            ? ValueProvider.Invoke(obj)
            : Property.GetValue(obj);

    }
}
