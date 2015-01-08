using System;

namespace noocyte.Utility.ExcelIO
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    internal sealed class IgnoreForExportAttribute : Attribute
    {
    }
}