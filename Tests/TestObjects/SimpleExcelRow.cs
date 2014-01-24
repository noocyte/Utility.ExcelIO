using System;
using System.Collections.Generic;

namespace Tests.TestObjects
{
    internal static class SimpleExcelRowFactory
    {
        internal static List<SimpleExcelRow> CreateSimpleRows()
        {
            var list = new List<SimpleExcelRow>();
            list.Add(new SimpleExcelRow() { Col1 = 1, Col2 = "A", Col3 = new DateTime(2012, 12, 12) });
            list.Add(new SimpleExcelRow() { Col1 = 2, Col2 = "B", Col3 = new DateTime(2012, 12, 12) });

            return list;
        }
    }


    internal class SimpleExcelRow
    {
        public int Col1 { get; set; }
        public string Col2 { get; set; }
        public DateTime Col3 { get; set; }
    }

}
