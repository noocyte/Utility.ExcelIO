using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;

namespace noocyte.Utility.ExcelIO
{

    internal class ExcelHeader
    {
        public string Text { get; set; }
        public bool Display { get; set; }
    }

    internal static class HeaderGenerator<T> where T : class
    {
        public static IEnumerable<ExcelHeader> GenerateHeaders()
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();

            foreach (PropertyInfo propertyInfo in propertyInfos)
            {
                var displayNameAttribute = propertyInfo.GetAttribute<DisplayNameAttribute>(false);
                var hasIgnoreAttribute = propertyInfo.GetAttribute<IgnoreForExportAttribute>(false);

                bool display = true;
                if (hasIgnoreAttribute != null)
                    display = false;

                string displayName = propertyInfo.Name;
                if (displayNameAttribute != null)
                    displayName = displayNameAttribute.DisplayName;

                yield return new ExcelHeader()
                {
                    Text = displayName,
                    Display = display
                };
            }
        }
    }
}