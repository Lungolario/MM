using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MMA.TestCases
{
    [TestClass]
    public class ExcelObject_Tests
    {
        [TestMethod()]
        public void ExcelObjectCreation()
        {
            ObjHandle obj = new Vol();
            object[,] range = new object[0, 0];
            obj.Create("C", range);
            Assert.AreEqual("C:0", obj.ToStringWithCounter());
            Assert.AreEqual("C", obj.ToString());
            obj.FinishMod();
            Assert.AreEqual("C:1", obj.ToStringWithCounter());
        }
    }
}
