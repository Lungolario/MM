using System;
using System.Collections.Generic;
using System.Linq;

namespace MMA
{
    public class Curve : ExcelObject
    {
        public string currency;
        public double? rate = null;
        public class RatesVectors : Vectors
        {
            public double[] start;
            public double[] end;
            public double[] contrate;
        }
        public RatesVectors rates = null;

        public override void CheckObject()
        {
            if (!(rate == null ^ rates == null))
                throw new Exception("Either field rate or table rates has to be set.");
        }

        private List<KeyValuePair<double, double>> timeLogDF { get; set; } = null;

        private void Bootstrap()
        {
            timeLogDF = new List<KeyValuePair<double, double>>();
            timeLogDF.Add(new KeyValuePair<double, double>(0, 0));
            for (int i = 0; i < rates.start.Length; i++)
            {
                if (rates.start[i] > timeLogDF.Max(kvp => kvp.Key))
                    throw new Exception("StartDate for instrument " + i + " after previous maximum enddate.");
                timeLogDF.Add(new KeyValuePair<double, double>(rates.end[i], GetLogDF(rates.start[i]) - rates.contrate[i] * (rates.end[i] - rates.start[i])));
            }
        }
        private double GetLogDF(double time)
        {
            if (timeLogDF == null)
                this.Bootstrap();
            if (timeLogDF.Exists(kvp => kvp.Key == time))
                return timeLogDF.First(kvp => kvp.Key == time).Value;
            KeyValuePair<double, double> before = timeLogDF.Where(kvp => kvp.Key < time).Last();
            KeyValuePair<double, double> after = timeLogDF.Where(kvp => kvp.Key > time).First();
            return ((time - before.Key) * after.Value + (after.Key - time) * before.Value) / (after.Key - before.Key);
        }
        public double GetDF(double time)
        {
            if (rate != null)
                return Math.Exp(- (double)rate * time);
            return Math.Exp(GetLogDF(time));
        }
        public double GetRate(double start, double end)
        {
            if (rate != null)
                return (double)rate;
            return (GetLogDF(start) - GetLogDF(end)) / (end - start);
        }
    }
}