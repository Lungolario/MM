using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace LungolarioMM
{
    static class CreateObj
    {
        public enum ObjType
        {
            MODEL,
            VOL,
            CURVE,
            RESULTS
        };
        public static List<string> Objs = new List<string>();
        static int ReturnCount(string name, ObjType type)
        {
            switch(type)
            {
                case ObjType.MODEL:
                    return CountNumberOfOccurences(name, ObjType.MODEL);
                case ObjType.VOL:
                    return CountNumberOfOccurences(name, ObjType.VOL);
                case ObjType.CURVE:
                    return CountNumberOfOccurences(name, ObjType.CURVE);
                case ObjType.RESULTS:
                    return CountNumberOfOccurences(name, ObjType.RESULTS);
                default:
                    return -1;
            }
        }
        public static int CountNumberOfOccurences(string name, ObjType type)
        {
            int counter = 0;
            foreach (string obj in Objs)
            {
                if (obj.ToString() == ConcatenateNameAndType(name, type.ToString()))
                    counter++;
            }
            AddObj(name,type.ToString());
            return counter;
        }
        static void AddObj(string name, string type)
        {
            Objs.Add(ConcatenateNameAndType(name , type));
        }
        static string ConcatenateNameAndType(string name, string type)
        {
            return name + "," + type;
        }
    }
    public class Lungolario
    {
        [ExcelFunction(Description = "Adds Hello World to the start of the string given as input")]
        public static string HelloWorld(string name)
        {
            return "Hello World " + name;
        }
        public static string mmCreateObj(string objName, string objType, object[,] range)
        {
            if (objType == "MODEL")
                return objName + ":" + CreateObj.CountNumberOfOccurences(objName, CreateObj.ObjType.MODEL);
            else if (objType == "VOL")
                return objName + ":" + CreateObj.CountNumberOfOccurences(objName, CreateObj.ObjType.VOL);
            else if (objType == "CURVE")
                return objName + ":" + CreateObj.CountNumberOfOccurences(objName, CreateObj.ObjType.CURVE);
            else if (objType == "RESULTS")
                return objName + ":" + CreateObj.CountNumberOfOccurences(objName, CreateObj.ObjType.RESULTS);
            else
                return "Error. No name/type found.";
        }
    }
}

