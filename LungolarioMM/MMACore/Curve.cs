using System;
using System.Collections;

namespace MMACore
{
    public class RatesVectors : Vectors
    {
        public double[] Start;
        public double[] End;
        public double[] ContRate;
    }
    public class Curve : ObjHandle
    {
        public string Currency;
        public double? Rate = null;

        public RatesVectors Rates = null;
        public enum BootStrapType {Normal, Reverse};
        public BootStrapType BootStrap;

        protected override void Check()
        {
            if (!(Rate == null ^ Rates == null))
                throw new Exception("Either field Rate or table Rates has to be set.");
        }

        private SortedList _timeLogDF;

        private void Bootstrap()
        {
            if (Rate != null)
            {
                _timeLogDF = new SortedList {{0.0, 0.0}, { 1000, -(double)Rate * 1000 } };
                return;
            }
            int i = 0;
            if (BootStrap == BootStrapType.Reverse)
                for (; i < Rates.Start.Length - 1; i++)
                    if (Rates.Start[i + 1].Equals(0) || Rates.Start[i + 1].Equals(Rates.End[i]))
                        break;
            if (i == 0)
                _timeLogDF = new SortedList { { 0.0, 0.0 } };
            else
            {
                _timeLogDF = new SortedList { { Rates.End[i], 0.0 } };
                for (int j = i; j >= 0; j--)
                    _timeLogDF.Add(Rates.Start[j], GetLogDF(Rates.End[j]) + Rates.ContRate[j] * (Rates.End[j] - Rates.Start[j]));
                double adjustToZero = GetLogDF(0);
                for (int j = 0; j < _timeLogDF.Count; j++)
                    _timeLogDF.SetByIndex(j, (double)_timeLogDF.GetByIndex(j) - adjustToZero);
                i++;
            }
            for (; i < Rates.Start.Length; i++)
                _timeLogDF.Add(Rates.End[i], GetLogDF(Rates.Start[i]) - Rates.ContRate[i] * (Rates.End[i] - Rates.Start[i]));
        }
        private double GetLogDF(double time)
        {
            if (_timeLogDF == null)
                Bootstrap();
            if (time > (double)_timeLogDF.GetKey(_timeLogDF.Count - 1))
                throw new ArgumentOutOfRangeException("Date " + time + " is after last date of bootstrapped curve!");
            if (time < (double)_timeLogDF.GetKey(0))
                throw new ArgumentOutOfRangeException("Date " + time + " is before first date during reverse bootstrapping of curve!");
            if (_timeLogDF.ContainsKey(time))
                return (double)_timeLogDF[time];
            int index = 0;
            for (; index < _timeLogDF.Count; index++)
                if ((double) _timeLogDF.GetKey(index) > time)
                    break;
            double beforeTime = (double)_timeLogDF.GetKey(index - 1);
            double afterTime = (double)_timeLogDF.GetKey(index);
            double beforeLogDF = (double)_timeLogDF.GetByIndex(index - 1);
            double afterLogDF = (double)_timeLogDF.GetByIndex(index);
            return ((time - beforeTime) * afterLogDF + (afterTime - time) * beforeLogDF) / (afterTime - beforeTime);
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