using System;
using ExcelDna.Integration;

namespace LungolarioMM
{
    public class ExcelFunctions
    {
        static ExcelObjectHandler objectHandler= new ExcelObjectHandler();
        ExcelFunctions()
        {
            objectHandler = new ExcelObjectHandler();
        }

        [ExcelFunction(Description = "Creates an object with name and type")]
        public static string mmCreateObj(string objName, string objType, object[,] range)
        {
            try
            {
                return objName + ":" + objectHandler.CreateObject(objName, objType.ToUpper(), range);
            }
            catch(Exception e)
            {
                return e.Message.ToString();
            }
        }

        [ExcelFunction(Description = "Display an object")]
        public static object[,] mmDisplayObj(string objName, string objType)
        {
            string[,] results;
            //Get the object
            ExcelObject dispObj = null;
            foreach(var obj in objectHandler.objList)
            {
                if (obj.name.ToUpper() == objName.ToUpper() && obj.GetType().Name.ToUpper() == objType.ToUpper())
                    dispObj = obj;
            }
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
            foreach(var obj in objectHandler.objList)
            {
                results[j, 0] = obj.name;
                results[j++, 1] = obj.GetType().Name.ToUpper();
            }
            return results;
        }
    }
}

