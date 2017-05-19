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
        [ExcelFunction(Description = "Delete all Objects")]
        public static void mmListObjs()
        {
            //Not implemented yet. Working on it.
            //int i = mmObjHandler.rangeForDisplay;
            //foreach(DNAObject obj in mmObjHandler.objs)
            //{
            //    ExcelReference cellOfName = new ExcelReference();
            //    cellOfName.SetValue(obj.name);
            //    ExcelReference cellOfType = new ExcelReference();
            //    cellOfType.SetValue(obj.name);
            //}
        }
    }
}

