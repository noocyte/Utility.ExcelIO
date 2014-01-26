using System;
using System.Collections.Generic;
using System.IO;
using noocyte.Utility.ExcelIO;
using NUnit.Framework;
using Tests.Helpers;
using Tests.TestObjects;

namespace Tests
{
    [TestFixture]
    public class GenericWriteExcelTest
    {
        [Test]
        public void CreateExcelFileTest()
        {
            const string filename = "SimpleBook2.xlsx";
            var target = new GenericWriteExcel();
            var rows = SimpleExcelRowFactory.CreateSimpleRows();

            var actual = target.CreateExcelFile(rows);
            using (var file = File.OpenWrite(filename))
            {
                file.Write(actual.GetBuffer(), 0, (int) actual.Position);
            }

            using (var file = File.OpenRead(filename))
            {
                var excel = new GenericReadExcel {HasHeader = true};
                var actualRows = excel.ExtractRows(file, row => new SimpleExcelRow
                {
                    Col1 = row.GetValue<int>(1),
                    Col2 = row.GetValue<string>(2),
                    Col3 = row.GetValue<DateTime>(3)
                }, 1);
                var actualRowsList = new List<SimpleExcelRow>();
                actualRowsList.AddRange(actualRows);

                AssertEx.PropertyValuesAreEquals(actualRowsList, rows);
            }
        }
    }
}