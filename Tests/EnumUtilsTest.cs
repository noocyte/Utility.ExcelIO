using System;
using noocyte.Utility.ExcelIO;
using NUnit.Framework;
using OfficeOpenXml.Table;

namespace Tests
{
    [TestFixture]
    public class EnumUtilsTest
    {
        public void ParseTestHelper<T>(string input = "Light9")
            where T : struct
        {
            const TableStyles expected = TableStyles.Light9;
            var actual = EnumUtils.Parse<T>(input);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [ExpectedException(typeof(ArgumentException))]
        public void Parse_InvalidTableStyle_ThrowsException_Test()
        {
            ParseTestHelper<TableStyles>("Not a valid input value");
        }

        [Test]
        [ExpectedException(typeof(ArgumentException))]
        public void Parse_NonEnum_ThrowsException_Test()
        {
            ParseTestHelper<EnumTestStruct>();
        }

        [Test]
        public void Parse_Proper_Test()
        {
            ParseTestHelper<TableStyles>();
        }
    }

    internal struct EnumTestStruct
    {
        public int Test;
    }
}