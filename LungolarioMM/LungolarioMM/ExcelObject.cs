using System;
using System.Collections.Generic;
using System.Reflection;

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
                System.Reflection.PropertyInfo[] keyList = this.GetType().GetProperties();
                for (int i = 0; i < range.GetLength(0); i++)
                    for (int j = 0; j < keyList.Length; j++)
                        if (range[i, 0].ToString().ToUpper() == keyList[j].Name.ToUpper())
                        {
                            var val = Convert.ChangeType(range[i, 1], keyList[j].PropertyType);
                            keyList[j].SetValue(this, val, null);
                            
                        }
            }
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
