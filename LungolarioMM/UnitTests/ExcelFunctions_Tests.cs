using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MMA.TestCases
{
    [TestClass]
    public class ExcelFunctions_Tests
    {
        [TestMethod()]
        public void TM_mmCreateObj_mmListObj()
        {
            string name = "A";
            string type = "MODEL";
            object[,] range = new object[0, 0];
            Assert.AreEqual(ExcelFunctions.mmCreateObj(name, type, range), "A:0");
            name = "B";
            type = "CURVE";
            Assert.AreEqual(ExcelFunctions.mmCreateObj(name, type, range), "B:0");
            object[,] result = ExcelFunctions.mmListObjs();
            Assert.AreEqual(result.GetLength(0), 2);
            Assert.AreEqual(result.GetLength(1), 2);
            Assert.AreEqual(result[0, 0].ToString(), "A");
            Assert.AreEqual(result[0, 1].ToString(), "MODEL");
            Assert.AreEqual(result[1, 0].ToString(), "B");
            Assert.AreEqual(result[1, 1].ToString(), "CURVE");
        }
    }
}