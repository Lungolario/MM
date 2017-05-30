using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MMA.TestCases
{
    [TestClass]
    public class ObjectCreationTest
    {
        [TestClass]
        public class ExcelObjectTests
        {
            [TestMethod()]
            public void ObjectCreationTestForNullValue()
            {
                ExcelObject obj = new Vol();
                object[,] rangeValues = new object[2, 2];
                rangeValues[0, 0] = "Currency";
                rangeValues[0, 1] = "Eur";
                rangeValues[1, 0] = "Volatility";
                rangeValues[1, 1] = "5.5";
                try
                {
                    obj.CreateObject("C", rangeValues);
                }
                catch (Exception e)
                {
                    StringAssert.Contains(e.Message, "not available for object");
                    return;
                }
                Assert.IsNotNull(obj);
            }
            [TestMethod()]
            public void ObjectCreationTestForType()
            {
                ExcelObject obj = new Vol();
                object[,] rangeValues = new object[2, 2];
                rangeValues[0, 0] = "currency";
                rangeValues[0, 1] = "Eur";
                rangeValues[1, 0] = "Volatility";
                rangeValues[1, 1] = "5.5";
                try
                {
                    obj.CreateObject("C", rangeValues);
                }
                catch (Exception e)
                {
                    StringAssert.Contains(e.Message, "not available for object");
                    return;
                }
                Assert.IsInstanceOfType(obj, typeof(Vol));
            }
        }
    }
}
