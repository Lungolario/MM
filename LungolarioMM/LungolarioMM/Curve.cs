using System;
using System.Collections.Generic;
using System.Linq;

namespace MMA
{
    public class Curve : ExcelObject
    {
        public string Currency;
        public double? Rate = null;
        public class RatesVectors : Vectors
        {
            public double[] Start;
            public double[] End;
            public double[] ContRate;
        }
        public RatesVectors Rates = null;

        public override void CheckObject()
        {
            if (!(Rate == null ^ Rates == null))
                throw new Exception("Either field Rate or table Rates has to be set.");
        }

        private List<KeyValuePair<double, double>> _timeLogDF;

        private void Bootstrap()
        {
            _timeLogDF = new List<KeyValuePair<double, double>>{ new KeyValuePair<double, double>(0, 0) };
            if (Rate != null)
                _timeLogDF.Add(new KeyValuePair<double, double>(1000, -(double)Rate * 1000));
            else
                for (int i = 0; i < Rates.Start.Length; i++)
                {
                    if (Rates.Start[i] > _timeLogDF.Max(kvp => kvp.Key))
                        throw new Exception("StartDate for instrument " + i + " after previous maximum enddate.");
                    _timeLogDF.Add(new KeyValuePair<double, double>(Rates.End[i], GetLogDF(Rates.Start[i]) - Rates.ContRate[i] * (Rates.End[i] - Rates.Start[i])));
                    _timeLogDF.Sort((x,y) => x.Key.CompareTo(y.Key));
                }
        }
        private double GetLogDF(double time)
        {
            if (_timeLogDF == null)
                Bootstrap();
            if (time > _timeLogDF.Max(kvp => kvp.Key))
                throw new Exception("Date " + time + " is after last date of bootstrapped curve!");
            if (_timeLogDF.Exists(kvp => kvp.Key == time))
                return _timeLogDF.First(kvp => kvp.Key == time).Value;
            KeyValuePair<double, double> before = _timeLogDF.Last(kvp => kvp.Key < time);
            KeyValuePair<double, double> after = _timeLogDF.First(kvp => kvp.Key > time);
            return ((time - before.Key) * after.Value + (after.Key - time) * before.Value) / (after.Key - before.Key);
        }

        public double GetDF(double time)
        {
            return Math.Exp(GetLogDF(time));
        }
        public double GetRate(double start, double end)
        {
            return (GetLogDF(start) - GetLogDF(end)) / (end - start);
        }
    }
}