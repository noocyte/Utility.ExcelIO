using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using noocyte.Utility.ExcelIO;
using Tests.Helpers;

namespace Tests
{
    
    
    /// <summary>
    ///This is a test class for ExtensionsTest and is intended
    ///to contain all ExtensionsTest Unit Tests
    ///</summary>
    [TestClass()]
    public class ExtensionsTest
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
        ///A test for GetAttribute
        ///</summary>
        public void GetAttributeTestHelper<T>(MemberInfo member, bool isRequired = true)
            where T : Attribute, new()
        {
            T expected = new T();
            T actual;
            actual = Extensions.GetAttribute<T>(member, isRequired);
            Assert.AreEqual(expected.GetType(), actual.GetType());
        }

        [TestMethod()]
        //[ExpectedException(typeof(ArgumentException))]
        public void GetAttribute_WithAttributePresent_Test()
        {
            MemberInfo member = typeof(ExtensionClass).GetProperty("TwoProp");
            GetAttributeTestHelper<DisplayNameAttribute>(member);
        }

        [TestMethod()]
        [ExpectedException(typeof(ArgumentException))]
        public void GetAttribute_WithOutAttributePresent_ThrowsException_Test()
        {
            MemberInfo member = typeof(ExtensionClass).GetProperty("OneProp");
            GetAttributeTestHelper<DisplayNameAttribute>(member);
        }

       

        [TestMethod()]
        public void GetPropertyValuesTest()
        {
            var currentDT = DateTime.Now;
            ExtensionClass obj = new ExtensionClass() { OneProp = currentDT, TwoProp = currentDT, ThreeProp = 3 };
            List<PropertyValueDefinition> expected = CreateValueList(currentDT);
            List<PropertyValueDefinition> actual = new List<PropertyValueDefinition>();
            actual.AddRange(Extensions.GetPropertyValues(obj));
            AssertEx.PropertyValuesAreEquals(actual, expected);
        }

        private List<PropertyValueDefinition> CreateValueList(object currentValue)
        {
            var list = new List<PropertyValueDefinition>();
            list.Add(new PropertyValueDefinition() { IsDateTime = true, PropertyValue = currentValue });
            list.Add(new PropertyValueDefinition() { IsDateTime = true, PropertyValue = currentValue });
            list.Add(new PropertyValueDefinition() { IsDateTime = false, PropertyValue = 3 });

            return list;
        }

        /// <summary>
        ///A test for IsDateTimeProperty
        /////</summary>
        //[TestMethod()]
        //public void IsDateTimePropertyTest()
        //{
        //    PropertyInfo propertyInfo = typeof(ExtensionClass).GetProperty("OneProp");
        //    bool expected = true;
        //    bool actual;
        //    actual = Extensions_Accessor.IsDateTimeProperty(propertyInfo);
        //    Assert.AreEqual(expected, actual);
        //}
    }

    class ExtensionClass
    {
        public DateTime OneProp { get; set; }
        [DisplayName("Second Property")]
        public DateTime TwoProp { get; set; }
        public int ThreeProp { get; set; }
        [IgnoreForExport]
        public int FourProp { get; set; }
    }
}
