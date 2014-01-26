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
    public class GenericReadExcelTest
    {
        public void ExtractRows_WithName_TestHelper()
        {
            var target = new GenericReadExcel {HasHeader = true};
            const string sheetName = "Sheet1";
            Func<ExcelWorksheetRow, SimpleExcelRow> rowFunction = ReadOneRow;
            var expected = SimpleExcelRowFactory.CreateSimpleRows();
            var actual = new List<SimpleExcelRow>();

            using (Stream stream = new FileStream("SimpleBook.xlsx", FileMode.Open))
            {
                actual.AddRange(target.ExtractRows(stream, rowFunction, sheetName));
            }

            AssertEx.PropertyValuesAreEquals(actual, expected);
        }

        public void ExtractRows_WithNumber_TestHelper()
        {
            var expected = SimpleExcelRowFactory.CreateSimpleRows();

            var target = new GenericReadExcel {HasHeader = true};
            var excelRows = new List<SimpleExcelRow>();

            using (Stream stream = new FileStream("SimpleBook.xlsx", FileMode.Open))
            {
                excelRows.AddRange(target.ExtractRows(stream,
                    row => new SimpleExcelRow
                    {
                        Col1 = row.GetValue<int>(1),
                        Col2 = row.GetValue<string>(2),
                        Col3 = row.GetValue<DateTime>(3)
                    },
                    sheetNumber: 1));
            }

            AssertEx.PropertyValuesAreEquals(excelRows, expected);
        }


        private static SimpleExcelRow ReadOneRow(ExcelWorksheetRow row)
        {
            return new SimpleExcelRow
            {
                Col1 = row.GetValue<int>(1),
                Col2 = row.GetValue<string>(2),
                Col3 = row.GetValue<DateTime>(3)
            };
        }

        [Test]
        public void ExtractRows_WithName_Test()
        {
            ExtractRows_WithName_TestHelper();
        }

        [Test]
        public void ExtractRows_WithNumber_Test()
        {
            ExtractRows_WithNumber_TestHelper();
        }
    }
}