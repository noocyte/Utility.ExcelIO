using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using noocyte.Utility.ExcelIO;
using OfficeOpenXml.Table;

namespace Tests
{
    /// <summary>
    ///This is a test class for EnumUtilsTest and is intended
    ///to contain all EnumUtilsTest Unit Tests
    ///</summary>
    [TestClass()]
    public class EnumUtilsTest
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


        /// <summary>
        ///A test for Parse
        ///</summary>
        public void ParseTestHelper<T>(string input = "Light9")
            where T : struct
        {
            TableStyles expected = TableStyles.Light9;
            T actual;
            actual = EnumUtils.Parse<T>(input);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        [ExpectedException(typeof(ArgumentException))]
        public void Parse_InvalidTableStyle_ThrowsException_Test()
        {
            ParseTestHelper<TableStyles>("Not a valid input value");
        }

        [TestMethod()]
        public void Parse_Proper_Test()
        {
            ParseTestHelper<TableStyles>();
        }

        [TestMethod()]
        [ExpectedException(typeof(ArgumentException))]
        public void Parse_NonEnum_ThrowsException_Test()
        {
            ParseTestHelper<EnumTestStruct>();
        }
    }

    struct EnumTestStruct
    {
        EnumTestStruct(int testVal)
        {
            Test = testVal;
        }
        public int Test;
    }
}
