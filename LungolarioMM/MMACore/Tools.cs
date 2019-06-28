using System;
using System.Reflection;

namespace MMACore
{
    public static class Tools
    {
        public static string StringTrim(string name)
        {
            var i = name.IndexOf(":", StringComparison.CurrentCulture);
            if (i > -1)
                return name.Substring(0, i);
            return name;
        }
    }

    //interface to simplify code in ObjHandle.CreateObject and ExcelFunctions.mmDisplayObj
    public interface IIMatrix
    {
        void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols);
        object[,] ObjInfo(object column, object row);
    }

    public class Matrix<T> : IIMatrix
    {
        internal T[,] Content;

        public virtual void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            Content = new T[nRows, nCols];
            for (var iRow = 0; iRow < nRows; iRow++)
                for (var iCol = 0; iCol < nCols; iCol++)
                    try
                    {
                        Content[iRow, iCol] = (T)range[rowStart + iRow, colStart + iCol];
                    }
                    catch (Exception e)
                    {
                        throw new Exception(e.Message + " Error in cell (" + (rowStart + iRow) + "," + (colStart + iCol) +
                                            ") of range.");
                    }
        }

        public virtual object[,] ObjInfo(object column, object row)
        {
            if (column.ToString().Length == 0 && row.ToString().Length == 0)
                return (object[,]) (object) Content;
            if (column.ToString().Length == 0)
            {
                var result = new object[1, Content.GetLength(1)];
                for (var i = 0; i < Content.GetLength(1); i++)
                    result[0, i] = Content[Convert.ToInt32(row), i];
                return result;
            }
            if (row.ToString().Length == 0)
            {
                var result = new object[Content.GetLength(0), 1];
                for (var i = 0; i < Content.GetLength(0); i++)
                    result[i, 0] = Content[i, Convert.ToInt32(column)];
                return result;
            }
            return new object[,] {{Content[Convert.ToInt32(row), Convert.ToInt32(column)]}};
        }
    }

    public class MatrixH<T, THEADER> : Matrix<T>
    {
        internal THEADER[] ColumnHeaders;

        public override void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            base.CreateMatrix(range, rowStart + 1, nRows - 1, colStart, nCols);
            ColumnHeaders = new THEADER[nCols];
            for (var iCol = 0; iCol < nCols; iCol++)
                try
                {
                    ColumnHeaders[iCol] = (THEADER) range[rowStart, colStart + iCol];
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message + " Error in cell (" + rowStart + "," + (colStart + iCol) +
                                        ") of range.");
                }
        }

        public override object[,] ObjInfo(object column, object row)
        {
            if (column.ToString().Length == 0 && row.ToString().Length == 0)
            {
                var result = new MatrixBuilder();
                result.Add(ColumnHeaders, false, true, true);
                result.Add(Content, false, true, false);
                return result.Deliver();
            }
            if (row.ToString().Length == 0 && Convert.ToInt32(row) == -1)
            {
                var result = new MatrixBuilder();
                result.Add(ColumnHeaders, false, true, true);
                return result.Deliver();
            }
            return base.ObjInfo(FindHeader(column, ColumnHeaders, "Column"), row);
        }

        public object FindHeader(object index, THEADER[] header, string name)
        {
            if (index.ToString().Length == 0)
                return index;
            for (var i = 0; i < header.GetLength(0); i++)
                if (index.Equals(header[i]))
                    return i;
            throw new Exception(name + " Header not found.");
        }
    }

    public class MatrixHR<T, THEADER> : MatrixH<T, THEADER>
    {
        internal THEADER[] RowHeaders;
        internal string UpperLeft;

        public override void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            base.CreateMatrix(range, rowStart, nRows, colStart + 1, nCols - 1);
            RowHeaders = new THEADER[nRows - 1];
            for (var iRow = 1; iRow < nRows; iRow++)
                try
                {
                    RowHeaders[iRow - 1] = (THEADER) range[rowStart + iRow, colStart];
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message + " Error in cell (" + (rowStart + iRow) + "," + colStart +
                                        ") of range.");
                }
            UpperLeft = range[rowStart, colStart].ToString();
        }

        public override object[,] ObjInfo(object column, object row)
        {
            if (column.ToString().Length == 0 && row.ToString().Length == 0)
            {
                var result = new MatrixBuilder();
                result.Add(new object[,] {{UpperLeft}}, false, true, false);
                result.Add(ColumnHeaders, true, false, true);
                result.Add(RowHeaders, false, true, false);
                result.Add(Content, true, false, false);
                return result.Deliver();
            }
            if (Convert.ToInt32(column) == -1 && Convert.ToInt32(row) == -1)
                return new object[,] {{UpperLeft}};
            if (Convert.ToInt32(column) == -1 && row.ToString().Length == 0)
            {
                var result = new object[Content.GetLength(0), 1];
                for (var iRow = 0; iRow < Content.GetLength(0); iRow++)
                    result[iRow, 0] = RowHeaders[iRow];
                return result;
            }
            return base.ObjInfo(column, FindHeader(row, RowHeaders, "Row"));
        }
    }

    public abstract class Vectors : IIMatrix
    {
        public void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            var keyList = GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
            if (nRows > 1 && nCols > 0)
                for (var i = 0; i < nCols; i++)
                {
                    int j;
                    for (j = 0; j < keyList.Length; j++)
                        if (range[rowStart, colStart + i].ToString().ToUpper() == keyList[j].Name.ToUpper())
                        {
                            if (range[rowStart + 1, colStart + i].ToString().Length == 0)
                            {
                                var help = Array.CreateInstance(keyList[j].FieldType.GetElementType(), 0);
                                keyList[j].SetValue(this, help);
                            }
                            else
                            {
                                var help = Array.CreateInstance(keyList[j].FieldType.GetElementType(), nRows - 1);
                                for (var k = 1; k < nRows; k++)
                                {
                                    var value = Convert.ChangeType(range[rowStart + k, colStart + i],
                                        keyList[j].FieldType.GetElementType());
                                    help.SetValue(value, k - 1);
                                }
                                keyList[j].SetValue(this, help);
                            }
                            break;
                        }
                    if (j == keyList.Length)
                        throw new Exception("Key " + range[rowStart, colStart + i] + " not available for table " +
                                            GetType().Name);
                }
        }

        public object[,] ObjInfo(object column, object row)
        {
            int iCol;
            var keyList = GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
            var nCol = keyList.Length;
            if (column.ToString().Length == 0 && row.ToString().Length == 0)
            {
                var header = new string[nCol];
                for (iCol = 0; iCol < nCol; iCol++)
                    header[iCol] = keyList[iCol].Name;
                var resi = new MatrixBuilder();
                resi.Add(header, false, true, true);
                resi.Add((Array) keyList[0].GetValue(this), false, true, false);
                for (iCol = 1; iCol < nCol; iCol++)
                    resi.Add((Array) keyList[iCol].GetValue(this), true, false, false);
                return resi.Deliver();
            }
            if (column.ToString().Length == 0 && Convert.ToInt32(row) == -1)
            {
                var result2 = new object[1, nCol];
                for (iCol = 0; iCol < nCol; iCol++)
                    result2[0, iCol] = keyList[iCol].Name;
                return result2;
            }
            for (iCol = 0; iCol < nCol; iCol++)
                if (keyList[iCol].Name.Equals(column))
                    break;
            if (iCol == nCol)
                return new object[,] {{"Column " + column + " not found!"}};
            var b = (Array) keyList[iCol].GetValue(this);
            if (b.Length == 0)
                return new object[,] {{"Column " + column + " has not been initialized!"}};
            if (row.ToString().Length == 0)
                return new[,] {{b.GetValue(Convert.ToInt32(row))}};
            var result = new object[b.Length, 1];
            for (int iRow = 0; iRow < b.Length; iRow++)
                result[iRow, 0] = b.GetValue(iRow);
            return result;
        }
    }
}