using System.Collections.Generic;
using System.Linq;
using noocyte.Utility.ExcelIO;
using NUnit.Framework;
using Tests.Helpers;
using Tests.TestObjects;

namespace Tests
{
    [TestFixture]
    public class HeaderGeneratorTest
    {
        internal void GenerateHeadersTestHelper<T>(IEnumerable<ExcelHeader> expected)
            where T : class
        {
            var actual = HeaderGenerator<T>.GenerateHeaders();
            var actualList = actual.ToList();
            AssertEx.PropertyValuesAreEquals(actualList, expected);
        }

        private static List<ExcelHeader> CreateHeaderForExcelRow()
        {
            var headerList = new List<ExcelHeader>
            {
                new ExcelHeader {Text = "Col1", Display = true},
                new ExcelHeader {Text = "Col2", Display = true},
                new ExcelHeader {Text = "Col3", Display = true}
            };

            return headerList;
        }

        [Test]
        public void GenerateAdvancedHeadersTest()
        {
            var expected = CreateHeaderForExcelRow();
            expected.Add(new ExcelHeader {Text = "HiddenCol", Display = false});
            expected.Add(new ExcelHeader {Text = "Last Col", Display = true});
            GenerateHeadersTestHelper<AdvancedExcelRow>(expected);
        }

        [Test]
        public void GenerateSimpleHeadersTest()
        {
            IEnumerable<ExcelHeader> expected = CreateHeaderForExcelRow();
            GenerateHeadersTestHelper<SimpleExcelRow>(expected);
        }
    }
}