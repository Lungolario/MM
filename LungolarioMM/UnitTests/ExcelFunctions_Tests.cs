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
            Assert.AreEqual(result.GetLength(0), 2);
            Assert.AreEqual(result.GetLength(1), 2);
            Assert.AreEqual(result[0, 0].ToString(), "A");
            Assert.AreEqual(result[0, 1].ToString(), "MODEL");
            Assert.AreEqual(result[1, 0].ToString(), "B");
            Assert.AreEqual(result[1, 1].ToString(), "VOL");

            Assert.AreEqual(ExcelFunctions.mmDeleteObjs("", "VOL"), "Deleted 1 object(s).");
        }
    }
}