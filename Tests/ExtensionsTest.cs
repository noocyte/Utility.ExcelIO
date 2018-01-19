using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using noocyte.Utility.ExcelIO;
using NUnit.Framework;
using Tests.Helpers;

namespace Tests
{
    [TestFixture]
    public class ExtensionsTest
    {
        public void GetAttributeTestHelper<T>(MemberInfo member, bool isRequired = true)
            where T : Attribute, new()
        {
            var expected = new T();
            var actual = member.GetAttribute<T>(isRequired);
            Assert.AreEqual(expected.GetType(), actual.GetType());
        }

        private static List<PropertyValueDefinition> CreateValueList(object currentValue)
        {
            var list = new List<PropertyValueDefinition>
            {
                new PropertyValueDefinition {IsDateTime = true, PropertyValue = currentValue},
                new PropertyValueDefinition {IsDateTime = true, PropertyValue = currentValue},
                new PropertyValueDefinition {IsDateTime = false, PropertyValue = 3}
            };

            return list;
        }

        [Test]
        public void GetAttribute_WithAttributePresent_Test()
        {
            MemberInfo member = typeof (ExtensionClass).GetProperty("TwoProp");
            GetAttributeTestHelper<DisplayNameAttribute>(member);
        }

        [Test]
        [ExpectedException(typeof(ArgumentException))]
        public void GetAttribute_WithOutAttributePresent_ThrowsException_Test()
        {
            MemberInfo member = typeof (ExtensionClass).GetProperty("OneProp");
            GetAttributeTestHelper<DisplayNameAttribute>(member);
        }


        [Test]
        public void GetPropertyValuesTest()
        {
            var currentDt = DateTime.Now;
            var obj = new ExtensionClass {OneProp = currentDt, TwoProp = currentDt, ThreeProp = 3};
            var expected = CreateValueList(currentDt);
            var actual = new List<PropertyValueDefinition>();
            actual.AddRange(obj.GetPropertyValues());
            AssertEx.PropertyValuesAreEquals(actual, expected);
        }

        [Test]
        public void IsDateTimePropertyTest()
        {
            var propertyInfo = typeof (ExtensionClass).GetProperty("OneProp");
            const bool expected = true;
            var actual = propertyInfo.IsDateTimeProperty();
            Assert.AreEqual(expected, actual);
        }
    }

    internal class ExtensionClass
    {
        public DateTime OneProp { get; set; }

        [DisplayName("Second Property")]
        public DateTime TwoProp { get; set; }

        public int ThreeProp { get; set; }

        [IgnoreForExport]
        public int FourProp { get; set; }
    }
}