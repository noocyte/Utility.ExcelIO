using System;
using System.ComponentModel;
using noocyte.Utility.ExcelIO;

namespace Tests.TestObjects
{
    internal class AdvancedExcelRow
    {
        public int Col1 { get; set; }
        public string Col2 { get; set; }        
        public DateTime Col3 { get; set; }
        [IgnoreForExport]
        public int HiddenCol { get; set; }
        [DisplayName("Last Col")]
        public string AdvancedCol { get; set; }
    }
}
