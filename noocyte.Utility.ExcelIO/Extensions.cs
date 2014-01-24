using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace noocyte.Utility.ExcelIO
{
    internal struct PropertyValueDefinition
    {
        public object PropertyValue { get; set; }
        public bool IsDateTime { get; set; }
    }

    internal static class Extensions
    {
        internal static IEnumerable<PropertyValueDefinition> GetPropertyValues<T>(this T obj) where T : class
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();

            foreach (PropertyInfo propertyInfo in propertyInfos)
            {
                var hasIgnoreAttribute = propertyInfo.GetCustomAttributes(typeof(IgnoreForExportAttribute), false);
                bool display = true;
                if (hasIgnoreAttribute != null && hasIgnoreAttribute.Count() > 0)
                    display = false;


                if (display)
                {
                    yield return new PropertyValueDefinition()
                    {
                        PropertyValue = propertyInfo.GetValue(obj, null),
                        IsDateTime = IsDateTimeProperty(propertyInfo),
                    };
                }
            }
        }

        private static bool IsDateTimeProperty(PropertyInfo propertyInfo)
        {
            return propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?);
        }

        internal static T GetAttribute<T>(this MemberInfo member, bool isRequired) where T : Attribute
        {
            var attribute = member.GetCustomAttributes(typeof(T), false).SingleOrDefault();

            if (attribute == null && isRequired)
            {
                throw new ArgumentException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "The {0} attribute must be defined on member {1}",
                        typeof(T).Name,
                        member.Name));
            }

            return (T)attribute;
        }
    }
}
