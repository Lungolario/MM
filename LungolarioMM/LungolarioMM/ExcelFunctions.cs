using System;
using ExcelDna.Integration;


namespace MMA
{
    public class ExcelFunctions
    {
        static ExcelObjectHandler objectHandler = new ExcelObjectHandler();
        ExcelFunctions()
        {
            objectHandler = new ExcelObjectHandler();
        }

        [ExcelFunction(Description = "Creates an object with name and type")]
        public static string mmCreateObj(string objName, string objType, object[,] range)
        {
            objName = Tools.StringTrim(objName);
            try
            {
                return objName + ":" + objectHandler.CreateObject(objName, objType.ToUpper(), range);
            }
            catch (Exception e)
            {
                return e.Message.ToString();
            }
        }

        [ExcelFunction(Description = "Display an object")]
        public static object[,] mmDisplayObj(string objName, string objType)
        {
            string[,] results;
            ExcelObject dispObj = objectHandler.GetObject(Tools.StringTrim(objName), objType);
            if (dispObj == null)
            {
                results = new string[1, 1];
                results[0, 0] = "Object not found.";
            }
            else
            {
                System.Reflection.PropertyInfo[] list = dispObj.GetType().GetProperties();
                results = new string[list.Length, 2];
                for (int i = 0; i < list.Length; i++)
                {
                    results[i, 0] = list[i].Name;
                    results[i, 1] = list[i].GetValue(dispObj, null).ToString();
                }
            }
            return results;
        }

        [ExcelFunction(Description = "Delete all objects")]
        public static string mmDeleteObjs()
        {
            int i = objectHandler.objList.Count;
            objectHandler.objList.Clear();
            return "Deleted " + i + " object(s).";
        }

        [ExcelFunction(Description = "List all objects")]
        public static object[,] mmListObjs()
        {
            string[,] results;
            int i = objectHandler.objList.Count;
            results = new string[i, 2];
            int j = 0;
            foreach (var obj in objectHandler.objList)
            {
                results[j, 0] = obj.name;
                results[j++, 1] = obj.GetType().Name.ToUpper();
            }
            return results;
        }

        [ExcelFunction(Description = "Get the instance of an object")]
        public static string mmGetObj(string objName, string objType)
        {
            objName = Tools.StringTrim(objName);
            ExcelObject obj = objectHandler.GetObject(objName, objType);
            if (obj != null)
                return objName.ToUpper() + ":" + obj.counter;
            else
                return "Object not found.";
        }

        [ExcelFunction(Description = "Modify an object")]
        public static string mmModifyObj(string objName, string objType, string key, object value)
        {
            ExcelObject obj = objectHandler.GetObject(Tools.StringTrim(objName), objType);
            if (obj == null)
                return "Object not found.";
            System.Reflection.PropertyInfo[] list = obj.GetType().GetProperties();
            for (int i = 0; i < list.Length; i++)
                if (key.ToUpper() == list[i].Name.ToUpper())
                {
                    try
                    {
                        list[i].SetValue(obj, value, null);
                        return objName.ToUpper() + ":" + ++obj.counter;
                    }
                    catch (Exception e)
                    {
                        return e.Message.ToString();
                    }
                }
            return "Object found, key not found.";
        }
    }
}

