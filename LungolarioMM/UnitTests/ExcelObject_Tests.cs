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
            Assert.AreEqual("C:0", obj.GetNameCounter());
            Assert.AreEqual("C", obj.GetName());
            obj.FinishMod();
            Assert.AreEqual("C:1", obj.GetNameCounter());
        }
    }
}
