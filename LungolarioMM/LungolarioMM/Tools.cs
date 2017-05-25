using System;

namespace MMA
{
    public static class Tools
    {
        public static string StringTrim(string name)
        {
            int i = name.IndexOf(":");
            if (i > -1)
                return name.Substring(0, i);
            return name;
        }
        public static T[] SubArray<T>(this T[] data, int index, int length)
        {
            T[] result = new T[length];
            Array.Copy(data, index, result, 0, length);
            return result;
        }
    }

    public class Matrix
    {
        public virtual void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            content = new object[nRows, nCols];
            for (int iRow = 0; iRow < nRows; iRow++)
                for (int iCol = 0; iCol < nCols; iCol++)
                    content[iRow, iCol] = range[rowStart + iRow, colStart + iCol];
        }
        public object[,] content;
        public bool hasHeader() { return false; }
        public bool hasRowHeader() { return false; }
        public int contentWidth() { return content.GetLength(1); }
        public int contentHeight() { return content.GetLength(0); }
    }
    public class MatrixH : Matrix
    {
        public override void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            base.CreateMatrix(range, rowStart + 1, nRows - 1, colStart, nCols);
            columnHeaders = new object[nCols];
            for (int iCol = 0; iCol < nCols; iCol++)
                columnHeaders[iCol] = range[rowStart, colStart + iCol];
        }
        public object[] columnHeaders;
        public new bool hasHeader() { return true; }
    }
    public class MatrixHR : MatrixH
    {
        public override void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            base.CreateMatrix(range, rowStart, nRows, colStart + 1, nCols - 1);
            rowHeaders = new object[nRows - 1];
            for (int iRow = 1; iRow < nRows; iRow++)
                rowHeaders[iRow] = range[rowStart + iRow, colStart];
            upperLeft = range[rowStart, colStart].ToString();
        }
        public object[] rowHeaders;
        public string upperLeft;
        public new bool hasRowHeader() { return true; }
    }
}