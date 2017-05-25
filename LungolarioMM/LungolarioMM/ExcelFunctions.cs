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
        public void AutoClose() {}

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
            string[,] results;
            ExcelObject dispObj = objectHandler.GetObject(objName, objType);
            if (dispObj == null)
            {
                results = new string[1, 1];
                results[0, 0] = "Object not found.";
            }
            else
            {
                PropertyInfo[] keyList = dispObj.GetType().GetProperties();
                results = new string[keyList.Length, 2];
                for (int i = 0; i < keyList.Length; i++)
                {
                    results[i, 0] = keyList[i].Name;
                    if (typeof(Matrix).IsAssignableFrom(keyList[i].PropertyType))
                        results[i, 1] = "Tables cannot be displayed yet";
                    else
                        results[i, 1] = keyList[i].GetValue(dispObj, null).ToString();
                }
            }
            return results;
        }

        [ExcelFunction(Description = "Delete objects")]
        public static string mmDeleteObjs(string name, string type)
        {
            int i = objectHandler.objList.Count;
            if (type != "" && name != "")
                objectHandler.objList.RemoveAll(item => item.name.ToUpper().Equals(name.ToUpper()) && item.GetType().Name.ToUpper() == type.ToUpper());
            else if (type != "")
                objectHandler.objList.RemoveAll(item=> item.GetType().Name.ToUpper() == type.ToUpper());
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
                results[j, 0] = obj.name;
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
        public static string mmGetObjInfo(string objName, string objType, string key)
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
                        return keyList[i].GetValue(obj, null).ToString();
                    }
                    catch (Exception e)
                    {
                        return e.Message.ToString();
                    }
                }
            return "Object found, key not found.";
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
                        return obj.IncreaseCounter().GetNameCounter();
                    }
                    catch (Exception e)
                    {
                        return e.Message.ToString();
                    }
                }
            return "Object found, key not found.";
        }

        [ExcelFunction(Description = "Save Objects to Txt File")]
        public static string mmSaveObjs(string objName, string objType, string location)
        {
            ExcelObject obj = objectHandler.GetObject(objName, objType);
            if (obj == null)
                return "No objects found for this name/type.";
            PropertyInfo[] keyList = obj.GetType().GetProperties();
            string data = "";
            data = obj.GetType().Name + "\r\n";
            data +="name " + obj.name + "\r\n";
            for (int i = 0; i < keyList.Length; i++)
            {
                data += keyList[i].Name + " " + keyList[i].GetValue(obj,null).ToString()+ "\r\n";
            }
            data += "\r\n";
            try
            {
                if (File.Exists(location))
                {
                    File.AppendAllText(location, data);
                }
                else
                {
                    File.WriteAllText(location, data);
                }
                return "1 object saved";
            }
            catch(Exception e)
            {
                return e.Message.ToString();
            }
        }

        [ExcelFunction(Description = "Load all Objects from Txt File")]
        public static object[,] mmLoadObjs(string location)
        {
            if (!File.Exists(location))
            {
                return null;
            }
            else
            {
                string[] text = File.ReadAllLines(location);
                int count = text.ToList().Where(item => item == "").Count();
                object[,] objNameAndCount = new object[count, 1];
                List<string> rangeValues = new List<string>();
                int counterForObjCount = 0;
                for (int counter = 0; counter < text.Length; counter++)
                {
                    if (counter % 5 == 0)
                    {
                        string type = text[counter];
                        int indexOfObjEnd = Array.IndexOf(text, "", counter++);
                        string name = text[counter ++].Replace("name ","");
                        rangeValues = Tools.SubArray(text, counter, indexOfObjEnd - counter).ToList();
                        object[,] range = new object[rangeValues.Count,2];
                        for(int i=0;i< rangeValues.Count;i++)
                        {
                            string[] str = rangeValues[i].Split(' ');
                            range[i, 0] = str[0];
                            range[i, 1] = str[1];
                        }
                        objNameAndCount[counterForObjCount, 0] = objectHandler.CreateObject(name, type, range).GetNameCounter();
                        counter = indexOfObjEnd;
                        counterForObjCount++;
                    }
                }
                return objNameAndCount;
            }
        }
    }
}

