using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineConnector
{
    internal class PropertyAttribute
    {
        public string FieldName { get; set; }
        public string ColumnName { get; set; }
        public Type FieldType { get; set; }
        public bool IsEmail { get; set; }
        public PropertyInfo Property { get; set; }

        public static IEnumerable<PropertyAttribute> GetPropertyAttributesForT(Type type)
        {
            var attributes = new List<PropertyAttribute>();
            foreach (var item in type.GetProperties())
            {
                var attr = item.GetCustomAttributes(typeof(SpPropertyAttribute), false)?.Cast<SpPropertyAttribute>()?.SingleOrDefault();
                if (attr?.Ignore == true)
                    continue;
                attributes.Add(new PropertyAttribute
                {
                    FieldName = item.Name,
                    ColumnName = attr?.FieldName ?? item.Name,
                    IsEmail = attr?.IsEmail ?? false,
                    FieldType = item.PropertyType,
                    Property = item
                });
            }
            return attributes;
        }
    }
}
