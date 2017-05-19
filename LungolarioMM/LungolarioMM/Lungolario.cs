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

        [ExcelFunction(Description = "Delete all Objects")]
        public static string mmDeleteObjs()
        {
            int i = objectHandler.objList.Count;
            objectHandler.objList.Clear();
            return "Deleted " + i + " object(s).";
        }

        [ExcelFunction(Description = "List all Objects")]
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

