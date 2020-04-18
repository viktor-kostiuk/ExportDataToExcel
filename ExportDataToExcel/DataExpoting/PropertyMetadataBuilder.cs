using System;
using System.Collections.Generic;
using System.Text;

namespace ExportDataToExcel.DataExpoting
{
    public class PropertyMetadataBuilder<TParent>
    {

        #region Properties

        public PropertyMetadata PropertyMetadata { get; }

        #endregion

        #region Constructors

        public PropertyMetadataBuilder(PropertyMetadata propertyMetadata)
        {
            PropertyMetadata = propertyMetadata ?? throw new ArgumentNullException(nameof(propertyMetadata));
        }

        #endregion

        #region Methods

        public PropertyMetadataBuilder<TParent> HasColumnName(string name)
        {
            PropertyMetadata.DisplayName = name ?? throw new ArgumentNullException(nameof(name));

            return this;
        }

        public PropertyMetadataBuilder<TParent> HasProvider(Func<TParent, object> converter)
        {
            PropertyMetadata.ValueProvider = (arg) => converter((TParent)arg);

            return this;
        }

        public PropertyMetadataBuilder<TParent> HasOrder(int order)
        {
            PropertyMetadata.Order = order;

            return this;
        }

        public PropertyMetadataBuilder<TParent> Ignore()
        {
            PropertyMetadata.IsIgnored = true;

            return this;
        }

        #endregion

    }
}
