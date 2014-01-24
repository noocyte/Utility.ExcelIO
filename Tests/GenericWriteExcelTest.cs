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
    ///This is a test class for GenericWriteExcelTest and is intended
    ///to contain all GenericWriteExcelTest Unit Tests
    ///</summary>
    [TestClass()]
    public class GenericWriteExcelTest
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


        [TestMethod()]
        public void CreateExcelFileTest()
        {
            string filename = "SimpleBook2.xlsx";
            GenericWriteExcel target = new GenericWriteExcel();
            List<SimpleExcelRow> rows = SimpleExcelRowFactory.CreateSimpleRows();

            MemoryStream actual;
            actual = target.CreateExcelFile(rows);
            using (FileStream file = File.OpenWrite(filename))
            {
                file.Write(actual.GetBuffer(), 0, (int)actual.Position);
            }

            using (FileStream file = File.OpenRead(filename))
            {
                var excel = new GenericReadExcel();
                excel.HasHeader = true;
                var actualRows = excel.ExtractRows<SimpleExcelRow>(file, delegate(ExcelWorksheetRow row)
                {
                    return new SimpleExcelRow()
                    {
                        Col1 = row.GetValue<int>(1),
                        Col2 = row.GetValue<string>(2),
                        Col3 = row.GetValue<DateTime>(3)
                    };
                }, 1);
                List<SimpleExcelRow> actualRowsList = new List<SimpleExcelRow>();
                actualRowsList.AddRange(actualRows);

                AssertEx.PropertyValuesAreEquals(actualRowsList, rows);
            }
           
        }
    }
}
