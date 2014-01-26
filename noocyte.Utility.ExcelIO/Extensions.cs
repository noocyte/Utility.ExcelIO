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
            var propertyInfos = typeof (T).GetProperties();

            return from propertyInfo in propertyInfos
                let hasIgnoreAttribute = propertyInfo.GetCustomAttributes(typeof (IgnoreForExportAttribute), false)
                let display = !(hasIgnoreAttribute.Any())
                where display
                select new PropertyValueDefinition
                {
                    PropertyValue = propertyInfo.GetValue(obj, null),
                    IsDateTime = IsDateTimeProperty(propertyInfo),
                };
        }

        internal static bool IsDateTimeProperty(this PropertyInfo propertyInfo)
        {
            return propertyInfo.PropertyType == typeof (DateTime) || propertyInfo.PropertyType == typeof (DateTime?);
        }

        internal static T GetAttribute<T>(this MemberInfo member, bool isRequired) where T : Attribute
        {
            var attribute = member.GetCustomAttributes(typeof (T), false).SingleOrDefault();

            if (attribute == null && isRequired)
            {
                throw new ArgumentException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "The {0} attribute must be defined on member {1}",
                        typeof (T).Name,
                        member.Name));
            }

            return (T) attribute;
        }
    }
}