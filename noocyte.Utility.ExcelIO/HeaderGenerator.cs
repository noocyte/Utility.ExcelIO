using System.Collections.Generic;
using System.ComponentModel;

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
            var propertyInfos = typeof (T).GetProperties();

            foreach (var propertyInfo in propertyInfos)
            {
                var displayNameAttribute = propertyInfo.GetAttribute<DisplayNameAttribute>(false);
                var hasIgnoreAttribute = propertyInfo.GetAttribute<IgnoreForExportAttribute>(false);

                var display = hasIgnoreAttribute == null;

                var displayName = propertyInfo.Name;
                if (displayNameAttribute != null)
                    displayName = displayNameAttribute.DisplayName;

                yield return new ExcelHeader
                {
                    Text = displayName,
                    Display = display
                };
            }
        }
    }
}