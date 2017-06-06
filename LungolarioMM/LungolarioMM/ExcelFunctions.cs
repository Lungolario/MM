using System;
using ExcelDna.Integration;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;

namespace MMA
{
    public class ExcelFunctions : IExcelAddIn
    {
        private static DateTime EXPIRY_DATE = new DateTime(2018, 05, 15);

        public void AutoOpen()
        {
            if (EXPIRY_DATE < DateTime.Today)
            {
                MessageBox.Show("The XLL has expired on " + EXPIRY_DATE.ToLongDateString(), "XLL expired");
                throw new Exception("XLL expired");
            }
        }
        public void AutoClose() { }

        static ExcelObjectHandler objectHandler = new ExcelObjectHandler();

        [ExcelFunction(Description = "Creates an object with name and type")]
        public static string mmCreateObj(string objName, string objType, object[,] range)
        {
            try
            {
                return objectHandler.CreateObject(objName, objType, range).GetNameCounter();
            }
            catch (Exception e)
            {
                return e.Message.ToString();
            }
        }

        [ExcelFunction(Description = "Display an object")]
        public static object[,] mmDisplayObj(string objName, string objType)
        {
            ExcelObject obj = objectHandler.GetObject(objName, objType);
            if (obj == null)
                return new string[1, 1] { { "Object not found." } };
            return obj.DisplayObject();
        }

        [ExcelFunction(Description = "Delete objects")]
        public static string mmDeleteObjs(string name, string type)
        {
            int i = objectHandler.objList.Count;
            if (type != "" && name != "")
                objectHandler.objList.RemoveAll(item => item.GetName().ToUpper().Equals(name.ToUpper()) && item.GetType().Name.ToUpper() == type.ToUpper());
            else if (type != "")
                objectHandler.objList.RemoveAll(item => item.GetType().Name.ToUpper() == type.ToUpper());
            else
                objectHandler.objList.Clear();
            i -= objectHandler.objList.Count;
            return "Deleted " + i + " object(s).";
        }

        [ExcelFunction(Description = "List all objects")]
        public static object[,] mmListObjs()
        {
            string[,] results = new string[objectHandler.objList.Count, 2];
            int j = 0;
            foreach (var obj in objectHandler.objList)
            {
                results[j, 0] = obj.GetName();
                results[j++, 1] = obj.GetType().Name.ToUpper();
            }
            return results;
        }

        [ExcelFunction(Description = "Get the instance of an object")]
        public static string mmGetObj(string objName, string objType)
        {
            ExcelObject obj = objectHandler.GetObject(objName, objType);
            if (obj == null)
                return "Object not found.";
            return obj.GetNameCounter();
        }

        [ExcelFunction(Description = "Get the value of a key of an object")]
        public static object[,] mmGetObjInfo(string objName, string objType, string key, object column, object row)
        {
            ExcelObject obj = objectHandler.GetObject(objName, objType);
            if (obj == null)
                return new string[1, 1] { { "Object not found." } };
            PropertyInfo[] keyList = obj.GetType().GetProperties();
            for (int i = 0; i < keyList.Length; i++)
                if (key.ToUpper() == keyList[i].Name.ToUpper())
                {
                    try
                    {
                        if (typeof(iMatrix).IsAssignableFrom(keyList[i].PropertyType))
                            return ((iMatrix)keyList[i].GetValue(obj, null)).ObjInfo(column, row);
                        return new string[1, 1] { { keyList[i].GetValue(obj, null).ToString() } };
                    }
                    catch (Exception e)
                    {
                        return new string[1, 1] { { e.Message.ToString() } };
                    }
                }
            return new string[1, 1] { { "Object found, key not found." } };
        }

        [ExcelFunction(Description = "Modify an object")]
        public static string mmModifyObj(string objName, string objType, string key, object value)
        {
            ExcelObject obj = objectHandler.GetObject(objName, objType);
            if (obj == null)
                return "Object not found.";
            PropertyInfo[] keyList = obj.GetType().GetProperties();
            for (int i = 0; i < keyList.Length; i++)
                if (key.ToUpper() == keyList[i].Name.ToUpper())
                {
                    try
                    {
                        keyList[i].SetValue(obj, value, null);
                        return obj.FinishMod().GetNameCounter();
                    }
                    catch (Exception e)
                    {
                        return e.Message.ToString();
                    }
                }
            return "Object found, key not found.";
        }

        [ExcelFunction(Description = "Save objects to file")]
        public static string mmSaveObjs(object[,] objNames, object[,] objTypes, string location)
        {
            string resultData = "";
            if (objNames.GetLength(0) != objTypes.GetLength(0) || objNames.GetLength(1) != 1 || objTypes.GetLength(1) != 1)
                return "objNames and objTypes must be Vectors of the same length!";
            for (int iObj = 0; iObj < objNames.GetLength(0); iObj++)
            {
                ExcelObject obj = objectHandler.GetObject(objNames[iObj, 0].ToString(), objTypes[iObj, 0].ToString());
                if (obj == null)
                    return "Object " + objNames[iObj, 0].ToString() + " of type " + objTypes[iObj, 0].ToString() + "does not exist!";
                resultData += obj.ObjectSerialize() + "\r\n";
            }
            try
            {
                File.WriteAllText(location, resultData);
            }
            catch (Exception e)
            {
                return e.Message.ToString();
            }
            return objNames.GetLength(0).ToString() + " object(s) was/were saved!";
        }

        [ExcelFunction(Description = "Load all objects from file")]
        public static object[,] mmLoadObjs(string location)
        {
            if (!File.Exists(location))
                return new string[1, 1] { { "File not found." } };
            string[] fileText = File.ReadAllLines(location);
            string name = "", type = "";
            List<string> loadedObjs = new List<string>();
            for (int iLine = 0, iObj = 0, iObjLine = 0; iLine < fileText.Length; iLine++, iObj++)
            {
                MatrixBuilder mRange = new MatrixBuilder();
                for (iObjLine = iLine; iObjLine < fileText.Length; iObjLine++)
                {
                    string[] lineFields = fileText[iObjLine].Split('\t');
                    if (iObjLine == iLine)
                    {
                        if (!lineFields[0].StartsWith("NEW"))
                            return new string[1, 1] { { "File not in correct format at line " + iObjLine } };
                        type = lineFields[0].Substring(3);
                    }
                    else if (iObjLine == iLine + 1)
                    {
                        if (!(lineFields[0] == "name"))
                            return new string[1, 1] { { "File not in correct format at line " + iObjLine } };
                        name = lineFields[1];
                    }
                    else if (lineFields.Length < 2)
                        break;
                    else
                        mRange.Add(lineFields, false, true, true);
                }
                try
                {
                    loadedObjs.Add(objectHandler.CreateObject(name, type, mRange.Deliver(false)).GetNameCounter().ToString());
                }
                catch (Exception e)
                {
                    return new string[1, 1] { { "Error when generating object " + name + " of type " + type + ". " + e.Message.ToString() } };
                }
                iLine = iObjLine;
            }
            object[,] vLoadedObjs = new object[loadedObjs.Count, 1];
            for (int i = 0; i < loadedObjs.Count; i++)
                vLoadedObjs[i, 0] = loadedObjs.ElementAt(i);

            return vLoadedObjs;
        }

        [ExcelFunction(Description = "Calculate rate from start to end")]
        public static object mmCurveRate(string objName, double start, double end)
        {
            ExcelObject curveObj = objectHandler.GetObject(objName, "CURVE");
            if (curveObj == null)
                return "Object not found.";
            try
            {
                return ((Curve)curveObj).GetRate(start, end);
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Calculate discount factor for time")]
        public static object mmCurveDF(string objName, double time)
        {
            ExcelObject curveObj = objectHandler.GetObject(objName, "CURVE");
            if (curveObj == null)
                return "Object not found.";
            try
            { 
                return ((Curve)curveObj).GetDF(time);
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
    }
}