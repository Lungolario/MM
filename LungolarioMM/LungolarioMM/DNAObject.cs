using System;
using System.Collections.Generic;
using System.Linq;

namespace LungolarioMM
{
    public abstract class ExcelObject
    {
        public string name;
        public int counter = 0;
        public virtual void CreateObject(string name, object[,] range)
        {
            this.name = name;
        }
    }
    public class Model : ExcelObject
    {
        public string modelname;
    }
    public class Curve : ExcelObject
    {
        public string currency;
        public double rate;
    }
    public class Vol : ExcelObject
    {
        public string currency;
        public double volatility;
    }
    public class Results : ExcelObject
    {
        public double result;
    }
    public class ExcelObjectHandler
    {
        public List<ExcelObject> objList = new List<ExcelObject>();
        public int CreateObject(string name, string type, object[,] range)
        {
            switch(type)
            {
                case "MODEL":
                    return CreateObject<Model>(name, range);
                case "CURVE":
                    return CreateObject<Curve>(name, range);
                case "VOL":
                    return CreateObject<Vol>(name, range);
                case "RESULTS":
                    return CreateObject<Results>(name, range);
                default:
                    throw new Exception("Cannot add this object. No type found!(" + type + ")");
            }
        }
        public int CreateObject<T>(string name, object[,] range) where T : ExcelObject, new()
        {
            ExcelObject newObj = new T();
            newObj.CreateObject(name, range);

            //handle if object with name and type already existed
            int index = -1;
            foreach (var existingObj in objList.OfType<T>())
            {
                if (existingObj.name == name)
                {
                    index = objList.IndexOf(existingObj);
                    newObj.counter = 1 + existingObj.counter;
                }
            }
            if (index > -1)
                objList.RemoveAt(index);

            objList.Add(newObj);
            return newObj.counter;
        }
    }
}
