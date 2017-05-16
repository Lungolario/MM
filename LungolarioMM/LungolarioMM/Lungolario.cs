using System;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

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
            return "Hello World " + name + DateTime.Now;
        }


        [ExcelFunction(Description = "Creates an object with name and type", IsMacroType = true)]
        public static string mmCreateObj(string objName, string objType, object[,] range)
        {
            string anyexception = "";
            mmObjHandler.rangeForDisplay = range;
            try
            {
                return objName + ":" + mmObjHandler.CreateObject(objName, objType.ToUpper());
            }
            catch (Exception e)
            {
                return anyexception = e.Message.ToString();
            }
        }


        [ExcelFunction(Description = "Delete all Objects")]
        public static void mmDeleteObj()
        {
            mmObjHandler.objs.Clear();
        }


        [ExcelFunction(Description = "Delete all Objects")]
        public static void mmListObj()
        {
            //int i = 0;
            //foreach (DNAObject obj in mmObjHandler.objs)
            //{
            //    ExcelReference cellOfName = new ExcelReference(mmObjHandler.startRow+i,mmObjHandler.startCol);
            //    cellOfName.SetValue(obj.name);
            //    ExcelReference cellOfType = new ExcelReference(mmObjHandler.startRow+i, mmObjHandler.startCol+1);
            //    cellOfType.SetValue(obj.ReturnType());
            //    i++;
            //}
        }
    }
}

