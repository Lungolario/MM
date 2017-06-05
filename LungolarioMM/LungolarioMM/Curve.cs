using System;
using System.Linq;

namespace MMA
{
    public class Curve : ExcelObject
    {
        public string currency { get; set; }
        public double? rate  { get; set; } = null; 
        public class RatesVectors : Vectors
        {
            public double[] start { get; set; }
            public double[] end { get; set; }
            public double[] contrate { get; set; }
        }
        public RatesVectors rates { get; set; } = null;

        public override void CheckObject()
        {
            if (!(rate == null ^ rates == null))
                throw new Exception("Either field rate or table rates has to be set.");
        }

        private double[] discDates = null;
        private double[] discFactorsLog = null;

        private void Bootstrap()
        {
            discDates = new double[rates.end.Length + 1];
            discFactorsLog = new double[rates.end.Length + 1];
            discDates[0] = 0;
            discFactorsLog[0] = 0;
            for (int i = 0; i < rates.end.Length; i++)
            {
                discDates[i + 1] = rates.end[i];
                double f = (new ArraySegment<double>(discDates, 0, i)).Where(n => n <= rates.start[i]).Max();
            }
        }

        public double GetRate(double start, double end)
        {
            if (rate != null)
                return (double)rate;
            this.Bootstrap();
            return 1;
        }
    }
}