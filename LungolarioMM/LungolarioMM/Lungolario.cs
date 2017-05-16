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
        public static string mmCreateObj(string objName, string objType, [ExcelArgument(AllowReference = false)]object[,] range)
        {
            string anyexception="";
            range[0, 0].ToString();

            try
            {
                mmObjHandler.CreateObject(objName, objType.ToUpper());
            }
            catch(Exception e)
            {
                anyexception = e.Message.ToString();
            }
            return mmObjHandler.CountAllObjects() + "\n" + anyexception + " ";
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

