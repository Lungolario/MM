using System;
using ExcelDna.Integration;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;

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
        public static object[,] CompileMatrix(List<string> data)
        {
            object[,] dataMatrix = null;
            int innerMatrixWidth = data[data.Count - 2].Split(' ').Count() - 1;
            int matrixWidth = data[data.Count - 2].Split(' ').Count();
            int matrixHeight = ((data[data.Count - 1].Split(' ').Count()-1) / (data[data.Count - 2].Split(' ').Count()-1)) + data.Count - 1;
            dataMatrix = new object[matrixHeight, matrixWidth];
            int rowCtr = 0;
            foreach (string entry in data)
            {
                string []dataValuePair = entry.Split(' ');
                for(int colCtr=0,valCtr=0;colCtr<dataValuePair.Length;colCtr++)
                {
                    if(rowCtr < data.Count-1)
                    {
                        if (dataValuePair[colCtr] != "")
                        {
                            if (rowCtr != data.Count - 2)
                                dataMatrix[rowCtr, colCtr] = dataValuePair[colCtr];
                            else
                            {
                                dataMatrix[rowCtr, ++colCtr] = dataValuePair[--colCtr];
                            }
                        }
                    }
                    else
                    {
                        if (dataValuePair[colCtr] != "")
                        {

                            dataMatrix[rowCtr, ++colCtr] = dataValuePair[valCtr];
                            valCtr++;
                            colCtr--;
                        }
                        if ((dataValuePair.Length - 1) / innerMatrixWidth == colCtr)
                        {
                            rowCtr++;
                            if (rowCtr > data.Count)
                                break;
                            colCtr = -1;
                        }

                    }
                }
                rowCtr++;
            }
            for(int rCtr=0; rCtr < dataMatrix.GetLength(0); rCtr++)
                for (int cCtr = 0; cCtr < dataMatrix.GetLength(1); cCtr++)
                    if (dataMatrix[rCtr,cCtr] == null)
                        dataMatrix[rCtr, cCtr] = (object)"#N/A";
            return dataMatrix;
        }
    }

    //interface to simplify code in ExcelObject.CreateObject and ExcelFunctions.mmDisplayObj
    public interface iMatrix
    {
        void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols);
        object[,] ObjInfo(object column, object row);
    }
    public class Matrix<T> : iMatrix
    {
        public virtual void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
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
        public virtual object[,] ObjInfo(object column, object row)
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
                    result[i, 0] = content[i, Convert.ToInt32(column)];
                return result;
            }
            return new object[1, 1] { { content[Convert.ToInt32(row), Convert.ToInt32(column)] } } ;
        }
        public T[,] content;
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
            if (column.GetType() == typeof(ExcelMissing) && row.GetType() == typeof(ExcelMissing))
            {
                MatrixBuilder result = new MatrixBuilder();
                result.add(columnHeaders, false, true, true);
                result.add(content, false, true, false);
                return result.deliver();
            }
            if (row.GetType() != typeof(ExcelMissing) && Convert.ToInt32(row) == -1)
            {
                MatrixBuilder result = new MatrixBuilder();
                result.add(columnHeaders, false, true, true);
                return result.deliver();
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
            if (column.GetType() == typeof(ExcelMissing) && row.GetType() == typeof(ExcelMissing))
            {
                MatrixBuilder result = new MatrixBuilder();
                result.add(new string[1, 1] { { upperLeft } },false,true,false);
                result.add(columnHeaders, true, false, true);
                result.add(rowHeaders, false, true, false);
                result.add(content, true, false, false);
                return result.deliver();
            }
            if (Convert.ToInt32(column) == -1 && Convert.ToInt32(row) == -1)
                return new string[1, 1] { { upperLeft } };
            if (Convert.ToInt32(column) == -1 && row.GetType() == typeof(ExcelMissing))
            {
                object[,] result = new object[content.GetLength(0), 1];
                for (int iRow = 0; iRow < content.GetLength(0); iRow++)
                    result[iRow, 0] = rowHeaders[iRow];
                return result;
            }
            return base.ObjInfo(column, FindHeader(row, rowHeaders, "Row"));
        }
        protected HR[] rowHeaders;
        protected string upperLeft;
    }

    public abstract class Vectors : iMatrix
    {
        public void CreateMatrix(object[,] range, int rowStart, int nRows, int colStart, int nCols)
        {
            PropertyInfo[] keyList = this.GetType().GetProperties();
            if (nRows > 1 && nCols > 0)
                for (int i = 0; i < nCols; i++)
                {
                    int j;
                    for (j = 0; j < keyList.Length; j++)
                        if (range[rowStart, colStart + i].ToString().ToUpper() == keyList[j].Name.ToUpper())
                        {
                            if (range[rowStart + 1, colStart + i] == ExcelEmpty.Value)
                            {
                                var help = Array.CreateInstance(keyList[j].PropertyType.GetElementType(), 0);
                                keyList[j].SetValue(this, help, null);
                            }
                            else
                            {
                                var help = Array.CreateInstance(keyList[j].PropertyType.GetElementType(), nRows - 1);
                                for (int k = 1; k < nRows; k++)
                                    help.SetValue(range[rowStart + k, colStart + i], k - 1);
                                keyList[j].SetValue(this, help, null);
                            }
                            break;
                        }
                    if (j == keyList.Length)
                        throw new Exception("Key " + range[i, 0].ToString() + " not available for table " + this.GetType().Name);
                }
        }
        public object[,] ObjInfo(object column, object row)
        {
            int iCol, iRow = 0;
            PropertyInfo[] keyList = this.GetType().GetProperties();
            int nCol = keyList.Length;
            if (column.GetType() == typeof(ExcelMissing) && row.GetType() == typeof(ExcelMissing))
            {
                string[] header = new string[nCol];
                for (iCol = 0; iCol < nCol; iCol++)
                    header[iCol] = keyList[iCol].Name;
                MatrixBuilder resi = new MatrixBuilder();
                resi.add(header, false, true, true);
                resi.add((Array)keyList[0].GetValue(this, null),false,true,false);
                for (iCol = 1; iCol < nCol; iCol++)
                    resi.add((Array)keyList[iCol].GetValue(this, null), true, false, false);
                return resi.deliver();
            }
            if (column.GetType() == typeof(ExcelMissing) && Convert.ToInt32(row) == -1)
            {
                object[,] result2 = new object[1, nCol];
                for (iCol = 0; iCol < nCol; iCol++)
                    result2[0, iCol] = keyList[iCol].Name;
                return result2;
            }
            for (iCol = 0; iCol < nCol; iCol++)
                if (keyList[iCol].Name.Equals(column))
                    break;
            if (iCol == nCol)
                return new string[1, 1] { { "Column " + column.ToString() + " not found!" } };
            var b = (Array)keyList[iCol].GetValue(this, null);
            if (b.Length == 0)
                return new string[1, 1] { { "Column " + column.ToString() + " has not been initialized!" } };
            if (row.GetType() != typeof(ExcelMissing))
                return new object[1, 1] { { b.GetValue(Convert.ToInt32(row))} };
            object[,] result = new object[b.Length, 1];
            for (iRow = 0; iRow < b.Length; iRow++)
                result[iRow, 0] = b.GetValue(iRow);
            return result;
        }
    }
}

