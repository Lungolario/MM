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

    //class to simplify code in ExcelObject.CreateObject
    public abstract class Mat
    {
        public abstract void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols);
    }
    public class Matrix<T> : Mat
    {
        public override void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            content = new T[nRows, nCols];
            for (int iRow = 0; iRow < nRows; iRow++)
                for (int iCol = 0; iCol < nCols; iCol++)
                {
                    try
                    {
                        content[iRow, iCol] = (T)range[rowStart + iRow, colStart + iCol];
                    }
                    catch (Exception e)
                    {
                        throw new Exception(e.Message.ToString() + " Error in cell (" + (rowStart+iRow) + "," + (colStart + iCol) + ") of range.");
                    }
                }
        }
        public T[,] content;
        public bool hasHeader() { return false; }
        public bool hasRowHeader() { return false; }
        public int contentWidth() { return content.GetLength(1); }
        public int contentHeight() { return content.GetLength(0); }
    }
    public class MatrixH<T, HR> : Matrix<T>
    {
        public override void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            base.CreateMatrix(range, rowStart + 1, nRows - 1, colStart, nCols);
            columnHeaders = new HR[nCols];
            for (int iCol = 0; iCol < nCols; iCol++)
            {
                try
                {
                    columnHeaders[iCol] = (HR)range[rowStart, colStart + iCol];
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message.ToString() + " Error in cell (" + rowStart + "," + (colStart + iCol) + ") of range.");
                }
            }
        }
        public HR[] columnHeaders;
        public new bool hasHeader() { return true; }
    }
    public class MatrixHR<T, HR> : MatrixH<T, HR>
    {
        public override void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            base.CreateMatrix(range, rowStart, nRows, colStart + 1, nCols - 1);
            rowHeaders = new HR[nRows - 1];
            for (int iRow = 1; iRow < nRows; iRow++)
            {
                try
                {
                    rowHeaders[iRow] = (HR)range[rowStart + iRow, colStart];
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message.ToString() + " Error in cell (" + (rowStart + iRow) + "," + colStart + ") of range.");
                }
            }
            upperLeft = range[rowStart, colStart].ToString();
        }
        public HR[] rowHeaders;
        public string upperLeft;
        public new bool hasRowHeader() { return true; }
    }
}