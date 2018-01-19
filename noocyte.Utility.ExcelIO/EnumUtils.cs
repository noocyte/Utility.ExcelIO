using System;
using System.Linq;

namespace noocyte.Utility.ExcelIO
{
    internal static class EnumUtils
    {
        internal static T Parse<T>(string input) where T : struct
        {
            if (!typeof (T).IsEnum)
                throw new ArgumentException("Generic Type 'T' must be an Enum");

            if (String.IsNullOrWhiteSpace(input))
                throw new ArgumentException("Unable to create enum value from input.");
            if (Enum.GetNames(typeof (T)).Any(
                e => String.Equals(e.Trim(), input.Trim(), StringComparison.InvariantCultureIgnoreCase)))
                return (T) Enum.Parse(typeof (T), input, true);
            throw new ArgumentException("Unable to create enum value from input.");
        }
    }
}