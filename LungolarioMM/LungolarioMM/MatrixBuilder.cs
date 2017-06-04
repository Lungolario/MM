using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace MMA
{
    public class MatrixBuilder
    {
        private List<MBHelp> matrices = new List<MBHelp>();
        private int upperLeftRow = -1;
        private int upperLeftColumn = -1;
        private int lowerRightRow = -1;
        private int lowerRightColumn = -1;
        private int maxColumn = -1;
        private int maxRow = -1;

        public MatrixBuilder Add(Array field, bool right, bool below, bool transpose)
        {
            if (below || maxRow < 0) upperLeftRow = maxRow + 1;
            upperLeftColumn = right ? lowerRightColumn + 1 : 0;
            if (field.Rank == 2)
            {
                lowerRightRow = upperLeftRow + field.GetLength(transpose ? 1 : 0) - 1;
                lowerRightColumn = upperLeftColumn + field.GetLength(transpose ? 0 : 1) - 1;
            }
            else
            {
                lowerRightRow = upperLeftRow + (transpose ? 1 : field.Length) - 1;
                lowerRightColumn = upperLeftColumn + (transpose ? field.Length : 1) - 1;
            }
            maxColumn = Math.Max(maxColumn, lowerRightColumn);
            maxRow = Math.Max(maxRow, lowerRightRow);
            matrices.Add(new MBHelp(field, upperLeftRow, upperLeftColumn, transpose));
            return this;
        }
        public object[,] Deliver(bool isForSaveDown = true)
        {
            object defaultValue;
            if (isForSaveDown)
                defaultValue = "";
            else
                defaultValue = ExcelEmpty.Value;
            object[,] result = new object[maxRow + 1, maxColumn + 1];
            for (int iRow = 0; iRow < maxRow + 1; iRow++)
                for (int iCol = 0; iCol < maxColumn + 1; iCol++)
                    result[iRow, iCol] = defaultValue;
            foreach (MBHelp field in matrices)
                field.Fill(ref result);
            return result;
        }
        private class MBHelp
        {
            private Array field;
            private int upperLeftRow;
            private int upperLeftColumn;
            private bool transpose;
            public MBHelp(Array field, int upperLeftRow, int upperLeftColumn, bool transpose)
            {
                this.field = field;
                this.upperLeftRow = upperLeftRow;
                this.upperLeftColumn = upperLeftColumn;
                this.transpose = transpose;
            }
            public void Fill(ref object[,] result)
            {
                if (field.Rank == 1)
                {
                    if (transpose)
                    {
                        for (int iCol = 0; iCol < field.Length; iCol++)
                            if (field.GetValue(iCol).ToString() != "")
                                result[upperLeftRow, upperLeftColumn + iCol] = field.GetValue(iCol);
                    }
                    else
                        for (int iRow = 0; iRow < field.Length; iRow++)
                            if (field.GetValue(iRow).ToString() != "")
                                result[upperLeftRow + iRow, upperLeftColumn] = field.GetValue(iRow);
                }
                else
                {
                    if (transpose)
                    {
                        for (int iCol = 0; iCol < field.GetLength(0); iCol++)
                            for (int iRow = 0; iRow < field.GetLength(1); iRow++)
                                if (field.GetValue(new int[2] { iCol, iRow }).ToString() != "")
                                    result[upperLeftRow + iRow, upperLeftColumn + iCol] = field.GetValue(new int[2] { iCol, iRow });
                    }
                    else
                        for (int iCol = 0; iCol < field.GetLength(1); iCol++)
                            for (int iRow = 0; iRow < field.GetLength(0); iRow++)
                                if (field.GetValue(new int[2] { iRow, iCol }).ToString() != "")
                                    result[upperLeftRow + iRow, upperLeftColumn + iCol] = field.GetValue(new int[2] { iRow, iCol });
                }
            }
        }
    }
}