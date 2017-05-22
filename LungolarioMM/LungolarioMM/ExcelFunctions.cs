using System;
using ExcelDna.Integration;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
namespace MMA
{
    public class ExcelFunctions : IExcelAddIn
    {
        static ExcelObjectHandler objectHandler = new ExcelObjectHandler();
        private static DateTime EXPIRY_DATE = new DateTime(2018, 05, 15);
        private static string expiredMessage = "Your Dll has Expired on " + EXPIRY_DATE.ToLongDateString();
        private static string expiryTitle = "Dll Expired";
        public void AutoOpen()
        {
            CheckDLL();
            objectHandler = new ExcelObjectHandler();
        }
        internal static bool Expired
        {
            get
            {
                if (EXPIRY_DATE < DateTime.Today)
                {
                    return true;
                }
                return false;
            }
        }
        public static void CheckDLL()
        {
            if (Expired)
            {
                MessageBox.Show(expiredMessage, expiryTitle);
                throw new Exception(expiredMessage);
            }
        }
        [ExcelFunction(Description = "Creates an object with name and type")]
        public static string mmCreateObj(string objName, string objType, object[,] range)
        {
            objName = Tools.StringTrim(objName);
            try
            {
                return objName + ":" + objectHandler.CreateObject(objName, objType, range);
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
                System.Reflection.PropertyInfo[] keyList = dispObj.GetType().GetProperties();
                results = new string[keyList.Length, 2];
                for (int i = 0; i < keyList.Length; i++)
                {
                    results[i, 0] = keyList[i].Name;
                    results[i, 1] = keyList[i].GetValue(dispObj, null).ToString();
                }
            }
            return results;
        }

        [ExcelFunction(Description = "Delete all objects")]
        public static string mmDeleteObjs()
        {
            int i = objectHandler.objList.Count;
            objectHandler.objList.Clear();
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
            if (obj != null)
                return obj.name + ":" + obj.counter;
            else
                return "Object not found.";
        }

        [ExcelFunction(Description = "Get the value of a key of an object")]
        public static string mmGetObjInfo(string objName, string objType, string key)
        {
            ExcelObject obj = objectHandler.GetObject(objName, objType);
            if (obj == null)
                return "Object not found.";
            System.Reflection.PropertyInfo[] keyList = obj.GetType().GetProperties();
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
            System.Reflection.PropertyInfo[] keyList = obj.GetType().GetProperties();
            for (int i = 0; i < keyList.Length; i++)
                if (key.ToUpper() == keyList[i].Name.ToUpper())
                {
                    try
                    {
                        keyList[i].SetValue(obj, value, null);
                        return obj.name + ":" + ++obj.counter;
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
            System.Reflection.PropertyInfo[] keyList = obj.GetType().GetProperties();
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
                List<object> objsLoadedWithNameCount = new List<object>();
                string[] text = File.ReadAllLines(location);
                string type, name;
                type = name = "";
                List<object[]> list = new List<object[]>();
                int counter = 0;
                foreach (string line in text)
                {
                    if (counter % 5 == 0)
                        type = line;
                    else if (counter % 5 == 1)
                    {
                        string[] str = line.Split(' ');
                        name = str[1];
                    }
                    else if (line == "")
                    {

                        object[,] range = new object[list.Count, 2];
                        for (int rangeCounter = 0; rangeCounter < list.Count; rangeCounter++)
                        {
                            range[rangeCounter, 0] = list[rangeCounter][0];
                            range[rangeCounter, 1] = list[rangeCounter][1];
                        }
                        objsLoadedWithNameCount.Add(mmCreateObj(name, type, range));
                        name = "";
                        type = "";
                        list.Clear();
                    }
                    else
                    {
                        string[] str = line.Split(' ');
                        object[] property = new string[2];
                        property[0] = (object)str[0];
                        property[1] = (object)str[1];
                        list.Add(property);
                    }
                    counter++;
                }
                string[,] nameCountOfObjs = new string[objsLoadedWithNameCount.Count, 1];
                int ctr = 0;
                foreach(object obj in objsLoadedWithNameCount)
                {
                    nameCountOfObjs[ctr++, 0] = obj.ToString();
                }
                return nameCountOfObjs;
            }
        }

        public void AutoClose()
        {
            throw new NotImplementedException();
        }
    }
}

