using System;
using System.Collections.Generic;
using System.Reflection;

namespace MMA
{
    public abstract class ExcelObject
    {
        public string name { get; private set; }
        private int counter = 0;
        public virtual void CreateObject(string name, object[,] range)
        {
            this.name = name;
            if (range.GetLength(0) > 0 && range.GetLength(1) > 1)
            {
                PropertyInfo[] keyList = this.GetType().GetProperties();
                for (int i = 0; i < range.GetLength(0); i++)
                {
                    int j;
                    for (j = 0; j < keyList.Length; j++)
                        if (range[i, 0].ToString().ToUpper() == keyList[j].Name.ToUpper())
                        {
                            if (typeof(Matrix).IsAssignableFrom(keyList[j].PropertyType))
                            {
                                int countRows, countColumns;
                                for (countRows = 1; countRows + i < range.GetLength(0); countRows++)
                                    if (range[countRows + i, 0].GetType() != typeof(ExcelDna.Integration.ExcelEmpty)) break;
                                for (countColumns = 1; countColumns < range.GetLength(1); countColumns++)
                                    if (range[i + 1, countColumns].GetType() == typeof(ExcelDna.Integration.ExcelEmpty)) break;

                                Matrix mat = (Matrix)Activator.CreateInstance(keyList[j].PropertyType);
                                mat.CreateMatrix(range, i + 1, countRows - 1, 1, countColumns - 1);
                                keyList[j].SetValue(this, mat, null);
                                i += countRows - 1;
                            }
                            else
                            {
                                var val = Convert.ChangeType(range[i, 1], keyList[j].PropertyType);
                                keyList[j].SetValue(this, val, null);
                            }
                            break;
                        }
                    if (j == keyList.Length)
                        throw new Exception("Key " + range[i, 0].ToString() + " not available for object " + this.GetType().ToString());
                }
            }
        }
        public ExcelObject TakeOverOldObject(ExcelObjectHandler objHandler)
        {
            ExcelObject existingObj = objHandler.GetObject(this.name, this.GetType().Name);
            if (existingObj != null)
            {
                this.counter = existingObj.counter + 1;
                objHandler.objList.Remove(existingObj);
            }
            return this;
        }
        public ExcelObject IncreaseCounter()
        {
            counter++;
            return this;
        }
        public string GetNameCounter()
        {
            return this.name + ":" + this.counter;
        }
    }
    public class Model : ExcelObject
    {
        public string modelName { get; set; }
        public int extraResults { get; set; }
    }
    public class Curve : ExcelObject
    {
        public string currency { get; set; }
        public double rate { get; set; }
        public MatrixH rates { get; set; }
    }
    public class Vol : ExcelObject
    {
        public string currency { get; set; }
        public double volatility { get; set; }
    }
    public class Results : ExcelObject
    {
        public double result { get; set; }
    }
    public class Security : ExcelObject
    {
        public string currency { get; set; }
        public double start { get; set; }
    }
    public class Dictionary : ExcelObject
    {
        public string MODE { get; set; }
        public string MODEL { get; set; }
        public string CURVE { get; set; }
        public string VOL { get; set; }
        public string SECURITY { get; set; }
        public string RESULTS { get; set; }
    }
    public class ExcelObjectHandler
    {
        public List<ExcelObject> objList = new List<ExcelObject>();
        public ExcelObject CreateObject(string name, string type, object[,] range)
        {
            ExcelObject newObj = (ExcelObject)Activator.CreateInstance(Type.GetType(typeof(ExcelObject).Namespace + "." + type, true, true));
            newObj.CreateObject(Tools.StringTrim(name), range);
            objList.Add(newObj.TakeOverOldObject(this));
            return newObj;
        }

        public ExcelObject GetObject(string name, string type)
        {
            return objList.Find(item => item.name.ToUpper().Equals(Tools.StringTrim(name).ToUpper()) && item.GetType().Name.ToUpper() == type.ToUpper());
        }
    }
}
