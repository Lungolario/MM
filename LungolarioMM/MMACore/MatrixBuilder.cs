using System;
using System.Collections.Generic;

namespace MMACore
{
    public class MatrixBuilder
    {
        private readonly List<MBHelp> _matrices = new List<MBHelp>();
        private int _upperLeftRow = -1;
        private int _upperLeftColumn = -1;
        private int _lowerRightRow = -1;
        private int _lowerRightColumn = -1;
        private int _maxColumn = -1;
        private int _maxRow = -1;

        public MatrixBuilder Add(Array field, bool right, bool below, bool transpose)
        {
            if (below || _maxRow < 0) _upperLeftRow = _maxRow + 1;
            _upperLeftColumn = right ? _lowerRightColumn + 1 : 0;
            if (field.Rank == 2)
            {
                _lowerRightRow = _upperLeftRow + field.GetLength(transpose ? 1 : 0) - 1;
                _lowerRightColumn = _upperLeftColumn + field.GetLength(transpose ? 0 : 1) - 1;
            }
            else
            {
                _lowerRightRow = _upperLeftRow + (transpose ? 1 : field.Length) - 1;
                _lowerRightColumn = _upperLeftColumn + (transpose ? field.Length : 1) - 1;
            }
            _maxColumn = Math.Max(_maxColumn, _lowerRightColumn);
            _maxRow = Math.Max(_maxRow, _lowerRightRow);
            _matrices.Add(new MBHelp(field, _upperLeftRow, _upperLeftColumn, transpose));
            return this;
        }
        public object[,] Deliver(object defaultValue = null)
        {
            if (defaultValue == null)
                defaultValue = "";
            object[,] result = new object[_maxRow + 1, _maxColumn + 1];
            for (int iRow = 0; iRow < _maxRow + 1; iRow++)
                for (int iCol = 0; iCol < _maxColumn + 1; iCol++)
                    result[iRow, iCol] = defaultValue;
            foreach (MBHelp field in _matrices)
                field.Fill(ref result);
            return result;
        }
        private class MBHelp
        {
            private readonly Array _field;
            private readonly int _upperLeftRow;
            private readonly int _upperLeftColumn;
            private readonly bool _transpose;
            public MBHelp(Array field, int upperLeftRow, int upperLeftColumn, bool transpose)
            {
                _field = field;
                _upperLeftRow = upperLeftRow;
                _upperLeftColumn = upperLeftColumn;
                _transpose = transpose;
            }
            public void Fill(ref object[,] result)
            {
                if (_field.Rank == 1)
                {
                    if (_transpose)
                    {
                        for (int iCol = 0; iCol < _field.Length; iCol++)
                            if (_field.GetValue(iCol).ToString().Length > 0)
                                result[_upperLeftRow, _upperLeftColumn + iCol] = _field.GetValue(iCol);
                    }
                    else
                        for (int iRow = 0; iRow < _field.Length; iRow++)
                            if (_field.GetValue(iRow).ToString().Length > 0)
                                result[_upperLeftRow + iRow, _upperLeftColumn] = _field.GetValue(iRow);
                }
                else
                {
                    if (_transpose)
                    {
                        for (int iCol = 0; iCol < _field.GetLength(0); iCol++)
                            for (int iRow = 0; iRow < _field.GetLength(1); iRow++)
                                if (_field.GetValue(new[] { iCol, iRow }).ToString().Length > 0)
                                    result[_upperLeftRow + iRow, _upperLeftColumn + iCol] = _field.GetValue(new[] { iCol, iRow });
                    }
                    else
                        for (int iCol = 0; iCol < _field.GetLength(1); iCol++)
                            for (int iRow = 0; iRow < _field.GetLength(0); iRow++)
                                if (_field.GetValue(new[] { iRow, iCol }).ToString().Length > 0)
                                    result[_upperLeftRow + iRow, _upperLeftColumn + iCol] = _field.GetValue(new[] { iRow, iCol });
                }
            }
        }
    }
}