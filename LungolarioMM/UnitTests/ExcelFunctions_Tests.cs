using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace MMA.TestCases
{
    [TestClass]
    public class ExcelFunctions_Tests
    {
        [TestMethod()]
        public void TM_mmCreateObj_mmListObj_mmDeletaObjs()
        {
            string name = "A";
            string type = "MODEL";
            object[,] range = new object[0, 0];
            Assert.AreEqual("A:0", ExcelFunctions.mmCreateObj(name, type, range));
            name = "B";
            type = "VOL";
            Assert.AreEqual("B:0", ExcelFunctions.mmCreateObj(name, type, range));
            object[,] result = ExcelFunctions.mmListObjs();
            Assert.AreEqual(2, result.GetLength(0));
            Assert.AreEqual(2, result.GetLength(1));
            Assert.AreEqual("A", result[0, 0].ToString());
            Assert.AreEqual("MODEL", result[0, 1].ToString());
            Assert.AreEqual("B", result[1, 0].ToString());
            Assert.AreEqual("VOL", result[1, 1].ToString());

            Assert.AreEqual("Deleted 1 object(s).", ExcelFunctions.mmDeleteObjs("", "VOL"));
        }
    }
}