using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace LungolarioMM
{
    public class Lungolario
    {
        static MMObjectHandler mmObjHandler= new MMObjectHandler();
        Lungolario()
        {
            mmObjHandler = new MMObjectHandler();
        }
        [ExcelFunction(Description = "Adds Hello World to the start of the string given as input")]
        public static string HelloWorld(string name)
        {
            return "Hello World " + name;
        }
        [ExcelFunction(Description = "Creates an object with name and type")]
        public static string mmCreateObj(string objName, string objType, object[,] range)
        {
            string anyexception="";
            mmObjHandler.rangeForDisplay = range;
            try
            {
                return objName + ":" + mmObjHandler.CreateObject(objName, objType.ToUpper());
            }
            catch(Exception e)
            {
                return anyexception = e.Message.ToString();
            }
        }

        [ExcelFunction(Description = "Delete all Objects")]
        public static string mmDeleteObjs()
        {
            int i = mmObjHandler.objs.Count;
            mmObjHandler.objs.Clear();
            return "Deleted " + i + " object(s).";
        }

        [ExcelFunction(Description = "List all Objects")]
        public static object[,] mmListObjs()
        {
            string[,] results;
            int i = mmObjHandler.objs.Count;
            results = new string[i, 2];
            int j = 0;
            foreach(var obj in mmObjHandler.objs)
            {
                results[j, 0] = obj.name;
                results[j++, 1] = obj.GetType().Name.ToUpper();
            }
            return results;
        }
    }
}

