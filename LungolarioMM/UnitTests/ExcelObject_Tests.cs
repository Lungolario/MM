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
            ExcelObject obj = new Vol();
            object[,] range = new object[0, 0];
            obj.CreateObject("C", range);
            Assert.AreEqual(obj.GetNameCounter(), "C:0");
            Assert.AreEqual(obj.GetName(), "C");
            obj.FinishMod();
            Assert.AreEqual(obj.GetNameCounter(), "C:1");
        }
    }
}
