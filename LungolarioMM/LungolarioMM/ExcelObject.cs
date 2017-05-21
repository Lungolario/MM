using System;
using System.Collections.Generic;

namespace MMA
{
    public abstract class ExcelObject
    {
        public string name;
        public int counter = 0;
        public virtual void CreateObject(string name, object[,] range)
        {
            this.name = name;
            if (range.GetLength(0) > 0 && range.GetLength(1) > 1)
            {
                System.Reflection.PropertyInfo[] list = this.GetType().GetProperties();
                for (int i = 0; i < range.GetLength(0); i++)
                    for (int j = 0; j < list.Length; j++)
                        if (range[i, 0].ToString().ToUpper() == list[j].Name.ToUpper())
                            list[j].SetValue(this, range[i, 1], null);
            }
        }
    }
    public class Model : ExcelObject
    {
        public string modelname { get; set; }
    }
    public class Curve : ExcelObject
    {
        public string currency { get; set; }
        public double rate { get; set; }
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
    public class ExcelObjectHandler
    {
        public List<ExcelObject> objList = new List<ExcelObject>();
        public int CreateObject(string name, string type, object[,] range)
        {
            ExcelObject newObj = (ExcelObject)Activator.CreateInstance(Type.GetType("MMA." + type, true, true));
            newObj.CreateObject(name, range);

            //handle if object with name and type already existed
            ExcelObject existingObj = GetObject(name, type);
            if (existingObj != null)
            {
                newObj.counter = 1 + existingObj.counter;
                objList.Remove(existingObj);
            }

            objList.Add(newObj);
            return newObj.counter;
        }

        public ExcelObject GetObject(string name, string type)
        {
            foreach (var existingObj in objList)
                if ((existingObj.name == name) && (existingObj.GetType().ToString().ToUpper() == "MMA." + type.ToUpper()))
                    return existingObj;
            return null;
        }
    }
}
