using System;
using System.Linq;

namespace noocyte.Utility.ExcelIO
{
    internal static class EnumUtils
    {
        internal static T Parse<T>(string input) where T : struct
        {
            //since we cant do a generic type constraint
            if (!typeof(T).IsEnum)
                throw new ArgumentException("Generic Type 'T' must be an Enum");

            if (!String.IsNullOrWhiteSpace(input))
            {
                if (Enum.GetNames(typeof(T)).Any(
                      e => e.Trim().ToUpperInvariant() == input.Trim().ToUpperInvariant()))
                {
                    return (T)Enum.Parse(typeof(T), input, true);
                }
            }
            throw new ArgumentException("Unable to create enum value from input.");
        }
    }
}
