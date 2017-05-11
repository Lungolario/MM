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
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            try
            {
                mmObjHandler.CreateObject(objName, objType.ToUpper());
            }
            catch(Exception e)
            {
                anyexception = e.Message.ToString();
            }
            return mmObjHandler.CountAllObjects() + anyexception;
        }
    }
}

