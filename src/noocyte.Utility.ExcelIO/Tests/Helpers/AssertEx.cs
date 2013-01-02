// Class from Juanma @Stackoverflow, modified by noocyte
// http://stackoverflow.com/questions/318210/compare-equality-between-two-objects-in-nunit 

using System.Collections;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests.Helpers
{
    internal static class AssertEx
    {

        internal static void PropertyValuesAreEquals(object actual, object expected)
        {
            // modification to handle pure lists as input
            if (actual is IList && expected is IList)
            {
                IList actualList = actual as IList;
                IList expectedList = expected as IList;
                AssertListsHaveSameNumberOfElements(actualList, expectedList);
                for (int i = 0; i < actualList.Count; i++)
                {
                    AssertEx.PropertyValuesAreEquals(actualList[i], expectedList[i]);
                }
            }
            else
            {

                PropertyInfo[] properties = expected.GetType().GetProperties();
                foreach (PropertyInfo property in properties)
                {
                    // need to handle indexer
                    var parameters = property.GetIndexParameters();

                    if (parameters.Length == 0)
                    {
                        object expectedValue = property.GetValue(expected, null);
                        object actualValue = property.GetValue(actual, null);

                        if (actualValue is IList)
                            AssertListsAreEquals(property, (IList)actualValue, (IList)expectedValue);
                        else if (!Equals(expectedValue, actualValue))
                            Assert.Fail("Property {0}.{1} does not match. Expected: {2} but was: {3}", property.DeclaringType.Name, property.Name, expectedValue, actualValue);
                    }
                }
            }
        }

        private static void AssertListsAreEquals(PropertyInfo property, IList actualList, IList expectedList)
        {
            AssertListsHaveSameNumberOfElements(actualList, expectedList);

            for (int i = 0; i < actualList.Count; i++)
                if (!Equals(actualList[i], expectedList[i]))
                    Assert.Fail("Property {0}.{1} does not match. Expected IList with element {1} equals to {2} but was IList with element {1} equals to {3}", property.PropertyType.Name, property.Name, expectedList[i], actualList[i]);
        }

        private static void AssertListsHaveSameNumberOfElements(IList actualList, IList expectedList)
        {
            if (actualList.Count != expectedList.Count)
                Assert.Fail("Expected IList containing {0} elements but was IList containing {1} elements", expectedList.Count, actualList.Count);
        }
    }
}
