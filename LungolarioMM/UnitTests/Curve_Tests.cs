using System;
using ExcelDna.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MMA.TestCases
{
    [TestClass]
    public class Curve_Tests
    {
        [TestMethod()]
        public void CurveBootstrapping()
        {
            object[,] range = new object[7, 4];
            for (int i = 0; i < range.GetLength(0); i++)
                for (int j = 0; j < range.GetLength(1); j++)
                    range[i, j] = ExcelEmpty.Value;
            range[0, 0] = "Currency";
            range[0, 1] = "USD";
            range[1, 0] = "Bootstrap";
            range[1, 1] = "Normal";
            range[2, 0] = "Rates";
            range[3, 1] = "Start";
            range[3, 2] = "End";
            range[3, 3] = "ContRate";
            range[4, 1] = 0;
            range[4, 2] = 1;
            range[4, 3] = 0.01;
            range[5, 1] = 0.5;
            range[5, 2] = 1.5;
            range[5, 3] = 0.02;
            range[6, 1] = 0;
            range[6, 2] = 2;
            range[6, 3] = 0.03;
            Curve obj = new Curve();
            obj.CreateObject("A", range);
            Assert.AreEqual(obj.GetRate(0, 0.5), 0.01);

            range[1, 1] = "Reverse";
            Curve obj2 = new Curve();
            obj2.CreateObject("B", range);
            Assert.AreEqual(obj2.GetRate(0, 0.5), 0);
            Assert.AreEqual(obj.GetDF(0), 1);

            var curveDisp = obj2.DisplayObject();
            Assert.AreEqual(curveDisp[6,0], "BootStrap");
        }
    }
}