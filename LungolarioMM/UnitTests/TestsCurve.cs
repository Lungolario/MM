using Microsoft.VisualStudio.TestTools.UnitTesting;
using MMACore;

namespace MMAExcel.TestCases
{
    [TestClass]
    public class TestsCurve
    {
        [TestMethod()]
        public void CurveBootstrapping()
        {
            object[,] range = new object[7, 4];
            for (int i = 0; i < range.GetLength(0); i++)
               for (int j = 0; j < range.GetLength(1); j++)
                    range[i, j] = "";
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
            obj.Create("A", range);
            Assert.AreEqual(0.01, obj.GetRate(0, 0.5));

            range[1, 1] = "Reverse";
            Curve obj2 = new Curve();
            obj2.Create("B", range);
            Assert.AreEqual(0, obj2.GetRate(0, 0.5));
            Assert.AreEqual(1, obj.GetDF(0));
            var curveDisp = obj2.Display();
            Assert.AreEqual("BootStrap", curveDisp[6,0]);
        }
    }
}