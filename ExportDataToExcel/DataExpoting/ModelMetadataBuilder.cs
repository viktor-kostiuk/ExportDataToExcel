using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace ExportDataToExcel.DataExpoting
{
    /// <summary>
    /// Provider a way to build ModelMetadata for specific type
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ModelMetadataBuilder<T>
    {

        private readonly Dictionary<string, PropertyMetadataBuilder<T>> _columnBuildersCache = new Dictionary<string, PropertyMetadataBuilder<T>>();

        public ModelMetadata Metadata { get; } = new ModelMetadata();

        public void InitFromProperties()
        {
            Type type = typeof(T);

            foreach (var p in type.GetProperties())
            {
                Metadata.Properties.Add(CreateForPropertyInfo(p));
            }
        }

        private PropertyMetadata CreateForPropertyInfo(PropertyInfo propertyInfo) => new PropertyMetadata
        {
            Property = propertyInfo,
            PropertyName = propertyInfo.Name,
            DisplayName = propertyInfo.Name
        };

        public PropertyMetadataBuilder<T> Property<TProperty>(Expression<Func<T, TProperty>> propertyExpression)
        {
            if (propertyExpression.Body is MemberExpression me)
            {
                return ResolvePropertyBuilder(me.Member.Name);
            }
            throw new ArgumentException("Invalid property expression", nameof(propertyExpression));
        }

        public PropertyMetadataBuilder<T> Property(string propertyName) =>
            ResolvePropertyBuilder(propertyName);

        private PropertyMetadataBuilder<T> ResolvePropertyBuilder(string propertyName)
        {
            if (!_columnBuildersCache.TryGetValue(propertyName, out PropertyMetadataBuilder<T> columnMetadataBuilder))
            {
                var propertyMetadata = Metadata.Properties.FirstOrDefault(x => x.PropertyName == propertyName);
                if (propertyMetadata == null)
                {
                    propertyMetadata = new PropertyMetadata
                    {
                        PropertyName = propertyName,
                        DisplayName = propertyName
                    };
                    Metadata.Properties.Add(propertyMetadata);
                }

                columnMetadataBuilder = new PropertyMetadataBuilder<T>(propertyMetadata);

                _columnBuildersCache.Add(propertyName, columnMetadataBuilder);
            }

            return columnMetadataBuilder;
        }

    }
}
