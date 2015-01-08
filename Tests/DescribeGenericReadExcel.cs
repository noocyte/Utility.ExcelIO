using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using noocyte.Utility.ExcelIO;
using NUnit.Framework;
using Tests.Helpers;
using Tests.TestObjects;

namespace Tests
{
    [TestFixture]
    public class DescribeGenericReadExcel
    {
        [Test]
        public void ItShouldExtractRowsAndHeader_GivenExcelFileWithMultipleSheets()
        {
            var target = new GenericReadExcel { HasHeader = true };
            IDictionary<string, IEnumerable<Dictionary<string, string>>> actual;

            using (Stream stream = new FileStream("ComplexBook.xlsx", FileMode.Open))
            {
                actual = target.ExtractRows(stream, (dictionary, row) =>
                {
                    var extractedRow = new Dictionary<string, string>();
                    foreach (var header in dictionary)
                    {
                        extractedRow[header.Value] = row.GetValue<string>(header.Key);
                    }

                    return extractedRow;

                });
            }

            actual.Count.Should().Be(2);
            actual.ContainsKey("col").Should().BeTrue();
            actual.ContainsKey("book").Should().BeTrue();
            actual["col"].Count().Should().Be(2);
            actual["book"].Last()["Name"].Should().Be("B");
        }


        [Test]
        public void ItShouldExtractRows_GivenSheetName()
        {
            var target = new GenericReadExcel { HasHeader = true };
            const string sheetName = "Sheet1";
            Func<Dictionary<int, string>, ExcelWorksheetRow, SimpleExcelRow> rowFunction = ReadOneRow;
            var expected = SimpleExcelRowFactory.CreateSimpleRows();
            var actual = new List<SimpleExcelRow>();

            using (Stream stream = new FileStream("SimpleBook.xlsx", FileMode.Open))
            {
                actual.AddRange(target.ExtractRows(stream, rowFunction, sheetName));
            }

            AssertEx.PropertyValuesAreEquals(actual, expected);
        }

        [Test]
        public void ItShouldExtractRows_GivenSheetNumber()
        {
            var expected = SimpleExcelRowFactory.CreateSimpleRows();

            var target = new GenericReadExcel { HasHeader = true };
            var excelRows = new List<SimpleExcelRow>();

            using (Stream stream = new FileStream("SimpleBook.xlsx", FileMode.Open))
            {
                excelRows.AddRange(target.ExtractRows(stream,
                    (header, row) => new SimpleExcelRow
                    {
                        Col1 = row.GetValue<int>(1),
                        Col2 = row.GetValue<string>(2),
                        Col3 = row.GetValue<DateTime>(3)
                    }, 1));
            }

            AssertEx.PropertyValuesAreEquals(excelRows, expected);
        }

        private static SimpleExcelRow ReadOneRow(Dictionary<int, string> header, ExcelWorksheetRow row)
        {
            return new SimpleExcelRow
            {
                Col1 = row.GetValue<int>(1),
                Col2 = row.GetValue<string>(2),
                Col3 = row.GetValue<DateTime>(3)
            };
        }
    }
}