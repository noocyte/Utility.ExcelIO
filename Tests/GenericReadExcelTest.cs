using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using noocyte.Utility.ExcelIO;
using Tests.Helpers;
using Tests.TestObjects;

namespace Tests
{


    /// <summary>
    ///This is a test class for GenericReadExcelTest and is intended
    ///to contain all GenericReadExcelTest Unit Tests
    ///</summary>
    [TestClass()]
    public class GenericReadExcelTest
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

        public void ExtractRows_WithName_TestHelper<T>()
        {
            GenericReadExcel target = new GenericReadExcel();
            target.HasHeader = true;
            string sheetName = "Sheet1";
            Func<ExcelWorksheetRow, SimpleExcelRow> rowFunction = ReadOneRow;
            List<SimpleExcelRow> expected = SimpleExcelRowFactory.CreateSimpleRows();
            List<SimpleExcelRow> actual = new List<SimpleExcelRow>();

            using (Stream stream = new FileStream("SimpleBook.xlsx", FileMode.Open))
            {
                actual.AddRange(target.ExtractRows<SimpleExcelRow>(stream, rowFunction, sheetName));
            }

            AssertEx.PropertyValuesAreEquals(actual, expected);
        }

        public void ExtractRows_WithNumber_TestHelper<T>()
        {
            var expected = SimpleExcelRowFactory.CreateSimpleRows();

            var target = new GenericReadExcel();
            target.HasHeader = true;
            var excelRows = new List<SimpleExcelRow>();

            using (Stream stream = new FileStream("SimpleBook.xlsx", FileMode.Open))
            {
                excelRows.AddRange(target.ExtractRows<SimpleExcelRow>(stream,
                    delegate(ExcelWorksheetRow row)
                    {
                        return new SimpleExcelRow()
                        {
                            Col1 = row.GetValue<int>(1),
                            Col2 = row.GetValue<string>(2),
                            Col3 = row.GetValue<DateTime>(3)
                        };
                    },
                    sheetNumber: 1));
            }

            AssertEx.PropertyValuesAreEquals(excelRows, expected);
        }

        

        private SimpleExcelRow ReadOneRow(ExcelWorksheetRow row)
        {
            return new SimpleExcelRow()
            {
                Col1 = row.GetValue<int>(1),
                Col2 = row.GetValue<string>(2),
                Col3 = row.GetValue<DateTime>(3)
            };
        }

        [TestMethod()]
        public void ExtractRows_WithName_Test()
        {
            ExtractRows_WithName_TestHelper<SimpleExcelRow>();
        }

        [TestMethod()]
        public void ExtractRows_WithNumber_Test()
        {
            ExtractRows_WithNumber_TestHelper<SimpleExcelRow>();
        }
    }
}
