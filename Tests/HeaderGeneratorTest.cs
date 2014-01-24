using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using noocyte.Utility.ExcelIO;
using Tests.Helpers;
using Tests.TestObjects;

namespace Tests
{


    /// <summary>
    ///This is a test class for HeaderGeneratorTest and is intended
    ///to contain all HeaderGeneratorTest Unit Tests
    ///</summary>
    [TestClass()]
    public class HeaderGeneratorTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        internal void GenerateHeadersTestHelper<T>(IEnumerable<ExcelHeader> expected)
            where T : class
        {
            IEnumerable<ExcelHeader> actual;
            actual = HeaderGenerator<T>.GenerateHeaders();
            var actualList = actual.ToList();
            AssertEx.PropertyValuesAreEquals(actualList, expected);
        }

        private List<ExcelHeader> CreateHeaderForExcelRow()
        {
            var headerList = new List<ExcelHeader>();
            headerList.Add(new ExcelHeader() { Text = "Col1", Display = true });
            headerList.Add(new ExcelHeader() { Text = "Col2", Display = true });
            headerList.Add(new ExcelHeader() { Text = "Col3", Display = true });

            return headerList;
        }

        [TestMethod()]
        public void GenerateSimpleHeadersTest()
        {
            IEnumerable<ExcelHeader> expected = CreateHeaderForExcelRow();
            GenerateHeadersTestHelper<SimpleExcelRow>(expected);
        }

        [TestMethod()]
        public void GenerateAdvancedHeadersTest()
        {
            var expected = CreateHeaderForExcelRow();
            expected.Add(new ExcelHeader() { Text = "HiddenCol", Display = false });
            expected.Add(new ExcelHeader() { Text = "Last Col", Display = true });
            GenerateHeadersTestHelper<AdvancedExcelRow>(expected);
        }
    }
}
