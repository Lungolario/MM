using System;
using ExcelDna.Integration;

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
        public abstract object[,] ObjInfo(object column, object row);
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
        public override object[,] ObjInfo(object column, object row)
        {
            if (column.GetType() == typeof(ExcelMissing) && row.GetType() == typeof(ExcelMissing))
                return (object[,])(object)content;
            if (column.GetType() == typeof(ExcelMissing))
            {
                object[,] result = new object[1, content.GetLength(1)];
                for (int i = 0; i < content.GetLength(1); i++)
                    result[0, i] = content[Convert.ToInt32(row), i];
                return result;
            }
            if (row.GetType() == typeof(ExcelMissing))
            {
                object[,] result = new object[content.GetLength(0), 1];
                for (int i = 0; i < content.GetLength(0); i++)
                {
                    result[i, 0] = content[i, Convert.ToInt32(column)];
                }
                return result;
            }
            return new object[1, 1] { { content[Convert.ToInt32(row), Convert.ToInt32(column)] } } ;
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
        public override object[,] ObjInfo(object column, object row)
        {
            int iCol;
            if (column.GetType() == typeof(ExcelMissing) && row.GetType() == typeof(ExcelMissing))
            {
                object[,] result = new object[1 + content.GetLength(0), content.GetLength(1)];
                for (iCol = 0; iCol < content.GetLength(1); iCol++)
                    result[0, iCol] = columnHeaders[iCol];
                for (int iRow = 0; iRow < content.GetLength(0); iRow++)
                    for (iCol = 0; iCol < content.GetLength(1); iCol++)
                        result[iRow + 1, iCol] = content[iRow, iCol];
                return result;
            }
            if (row.GetType() != typeof(ExcelMissing) && Convert.ToInt32(row) == 0)
            {
                object[,] result = new object[1, content.GetLength(1)];
                for (int i = 0; i < content.GetLength(1); i++)
                    result[0, i] = columnHeaders[i];
                return result;
            }
            return base.ObjInfo(FindHeader(column, columnHeaders, "Column"), row);
        }
        public object FindHeader(object index, HR[] header, string name)
        {
            if (index.GetType() == typeof(ExcelMissing))
                return index;
            for (int i = 0; i < header.GetLength(0); i++)
                if (index.Equals((object)header[i]))
                    return i;
            throw new Exception(name + " Header not found.");
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
                    rowHeaders[iRow - 1] = (HR)range[rowStart + iRow, colStart];
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message.ToString() + " Error in cell (" + (rowStart + iRow) + "," + colStart + ") of range.");
                }
            }
            upperLeft = range[rowStart, colStart].ToString();
        }
        public override object[,] ObjInfo(object column, object row)
        {
            int iCol, iRow;
            if (column.GetType() == typeof(ExcelMissing) && row.GetType() == typeof(ExcelMissing))
            {
                object[,] result = new object[1 + content.GetLength(0), 1 + content.GetLength(1)];
                result[0, 0] = upperLeft;
                for (iRow = 0; iRow < content.GetLength(0); iRow++)
                    result[iRow + 1, 0] = rowHeaders[iRow];
                for (iCol = 0; iCol < content.GetLength(1); iCol++)
                    result[0, iCol + 1] = columnHeaders[iCol];
                for (iRow = 0; iRow < content.GetLength(0); iRow++)
                    for (iCol = 0; iCol < content.GetLength(1); iCol++)
                        result[iRow + 1, iCol] = content[iRow, iCol];
                return result;
            }
            if (Convert.ToInt32(column) == 0 && Convert.ToInt32(row) == 0)
                return new string[1, 1] { { upperLeft } };
            return base.ObjInfo(column, FindHeader(row, rowHeaders, "Row"));
        }
        public HR[] rowHeaders;
        public string upperLeft;
        public new bool hasRowHeader() { return true; }
    }
}