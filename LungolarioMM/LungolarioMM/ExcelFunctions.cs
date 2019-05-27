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
        private static DateTime _expiryDate = new DateTime(2020, 05, 15);

        public void AutoOpen()
        {
            if (_expiryDate < DateTime.Today)
            {
                MessageBox.Show("The XLL has expired on " + _expiryDate.ToLongDateString(), "XLL expired");
                throw new Exception("XLL expired");
            }
        }
        public void AutoClose() { }

        static readonly ExcelObjectHandler ObjectHandler = new ExcelObjectHandler();

        [ExcelFunction(Description = "Creates an object with name and type", IsVolatile = true)]
        public static string mmCreateObj(string objName, string objType, object[,] range)
        {
            try
            {
                return ObjectHandler.CreateOrOverwriteObject(objName, objType, range).GetNameCounter();
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Display an object", IsVolatile = true)]
        public static object[,] mmDisplayObj(string objName, string objType)
        {
            ExcelObject obj = ObjectHandler.GetObject(objName, objType);
            if (obj == null)
                return new object[,] { { "Object not found." } };
            return obj.DisplayObject();
        }

        [ExcelFunction(Description = "Delete objects", IsVolatile = true)]
        public static string mmDeleteObjs(string name, string type)
        {
            return "Deleted " + ObjectHandler.DeleteObjects(name, type) + " object(s).";
        }

        [ExcelFunction(Description = "List all objects", IsVolatile = true)]
        public static object[,] mmListObjs()
        {
            object[,] results = new object[ObjectHandler.ObjList.Count, 2];
            int j = 0;
            foreach (var obj in ObjectHandler.ObjList)
            {
                results[j, 0] = obj.GetName();
                results[j++, 1] = obj.GetType().Name.ToUpper();
            }
            return results;
        }

        [ExcelFunction(Description = "Get the instance of an object", IsVolatile = true)]
        public static string mmGetObj(string objName, string objType)
        {
            ExcelObject obj = ObjectHandler.GetObject(objName, objType);
            if (obj == null)
                return "Object not found.";
            return obj.GetNameCounter();
        }

        [ExcelFunction(Description = "Get the value of a key of an object", IsVolatile = true)]
        public static object[,] mmGetObjInfo(string objName, string objType, string key, object column, object row)
        {
            ExcelObject obj = ObjectHandler.GetObject(objName, objType);
            if (obj == null)
                return new object[,] { { "Object not found." } };
            FieldInfo[] keyList = obj.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
            foreach (var keyL in keyList)
                if (key.ToUpper() == keyL.Name.ToUpper())
                {
                    try
                    {
                        if (typeof(IIMatrix).IsAssignableFrom(keyL.FieldType))
                            return ((IIMatrix)keyL.GetValue(obj)).ObjInfo(column, row);
                        return new object[,] { { keyL.GetValue(obj).ToString() } };
                    }
                    catch (Exception e)
                    {
                        return new object[,] { { e.Message } };
                    }
                }
            return new object[,] { { "Object found, key not found." } };
        }

        [ExcelFunction(Description = "Modify an object", IsVolatile = true)]
        public static string mmModifyObj(string objName, string objType, string key, object value)
        {
            ExcelObject obj = ObjectHandler.GetObject(objName, objType);
            if (obj == null)
                return "Object not found.";
            FieldInfo[] keyList = obj.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
            foreach (var keyL in keyList)
                if (key.ToUpper() == keyL.Name.ToUpper())
                {
                    try
                    {
                        keyL.SetValue(obj, value);
                        return obj.FinishMod().GetNameCounter();
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }
                }
            return "Object found, key not found.";
        }

        [ExcelFunction(Description = "Save objects to file", IsVolatile = true)]
        public static string mmSaveObjs(object[,] objNames, object[,] objTypes, string location)
        {
            string resultData = "";
            if (objNames.GetLength(0) != objTypes.GetLength(0) || objNames.GetLength(1) != 1 || objTypes.GetLength(1) != 1)
                return "objNames and objTypes must be Vectors of the same length!";
            for (int iObj = 0; iObj < objNames.GetLength(0); iObj++)
            {
                ExcelObject obj = ObjectHandler.GetObject(objNames[iObj, 0].ToString(), objTypes[iObj, 0].ToString());
                if (obj == null)
                    return "Object " + objNames[iObj, 0] + " of type " + objTypes[iObj, 0] + "does not exist!";
                resultData += obj.ObjectSerialize() + "\r\n";
            }
            try
            {
                File.WriteAllText(location, resultData);
            }
            catch (Exception e)
            {
                return e.Message;
            }
            return objNames.GetLength(0) + " object(s) was/were saved!";
        }

        [ExcelFunction(Description = "Load all objects from file", IsVolatile = true)]
        public static object[,] mmLoadObjs(string location)
        {
            if (!File.Exists(location))
                return new object[,] { { "File not found." } };
            string[] fileText = File.ReadAllLines(location);
            string name = "", type = "";
            List<string> loadedObjs = new List<string>();
            for (int iLine = 0; iLine < fileText.Length; iLine++)
            {
                int iObjLine;
                MatrixBuilder mRange = new MatrixBuilder();
                for (iObjLine = iLine; iObjLine < fileText.Length; iObjLine++)
                {
                    string[] lineFields = fileText[iObjLine].Split('\t');
                    if (iObjLine == iLine)
                    {
                        if (!lineFields[0].StartsWith("NEW"))
                            return new object[,] { { "File not in correct format at line " + iObjLine } };
                        type = lineFields[0].Substring(3);
                    }
                    else if (iObjLine == iLine + 1)
                    {
                        if (lineFields[0] != "name")
                            return new object[,] { { "File not in correct format at line " + iObjLine } };
                        name = lineFields[1];
                    }
                    else if (lineFields.Length < 2)
                        break;
                    else
                        mRange.Add(lineFields, false, true, true);
                }
                try
                {
                    loadedObjs.Add(ObjectHandler.CreateOrOverwriteObject(name, type, mRange.Deliver(false)).GetNameCounter());
                }
                catch (Exception e)
                {
                    return new object[,] { { "Error when generating object " + name + " of type " + type + ". " + e.Message} };
                }
                iLine = iObjLine;
            }
            object[,] vLoadedObjs = new object[loadedObjs.Count, 1];
            for (int i = 0; i < loadedObjs.Count; i++)
                vLoadedObjs[i, 0] = loadedObjs.ElementAt(i);

            return vLoadedObjs;
        }

        [ExcelFunction(Description = "Calculate rate from Start to End", IsVolatile = true)]
        public static object mmCurveRate(string objName, double start, double end)
        {
            ExcelObject curveObj = ObjectHandler.GetObject(objName, "CURVE");
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

        [ExcelFunction(Description = "Calculate discount factor for time", IsVolatile = true)]
        public static object mmCurveDF(string objName, double time)
        {
            ExcelObject curveObj = ObjectHandler.GetObject(objName, "CURVE");
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