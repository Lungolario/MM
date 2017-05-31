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

        private static object[,] GetObjInfoAs2D(ExcelObject obj)
        {
            PropertyInfo[] keyList = obj.GetType().GetProperties();
            MatrixBuilder result = new MatrixBuilder();
            for (int i = 0; i < keyList.Length; i++)
            {
                if (keyList[i].GetValue(obj, null) != null)
                {
                    result.Add(new string[1] { keyList[i].Name }, false, true, false);
                    if (typeof(iMatrix).IsAssignableFrom(keyList[i].PropertyType))
                        result.Add(((iMatrix)keyList[i].GetValue(obj, null)).ObjInfo(ExcelMissing.Value, ExcelMissing.Value), true, true, false);
                    else
                        result.Add(new object[1] { keyList[i].GetValue(obj, null) }, true, false, false);
                }
            }
            return result.Deliver();
        }



        /***************************
         * Excel Functions
         * ************************/
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
            return GetObjInfoAs2D(obj);
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
                        {
                            return ((iMatrix) keyList[i].GetValue(obj, null)).ObjInfo(column, row);
                        }
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
        public static string mmSaveObjs(object[,] objName, object[,] objType, string location)
        {
            string data = "";
            ExcelObject obj;
            for (int ctr=0;ctr<objName.Length;ctr++)
            {
                obj = objectHandler.GetObject(objName[ctr, 0].ToString(), objType[ctr, 0].ToString());
                if (obj != null)
                {
                    object [,] objData = GetObjInfoAs2D(obj);
                    data += obj.GetType().Name + "\r\n";
                    data += "name " + obj.GetName() + "\r\n";
                    for (int rCtr=0;rCtr<objData.GetLength(0);rCtr++)
                    {
                        for(int cCtr=0;cCtr<objData.GetLength(1);cCtr++)
                        {
                            if(objData[rCtr,cCtr].ToString()!="")
                            {
                                data += objData[rCtr,cCtr].ToString() + " ";
                            }
                        }
                        data += "\r\n";
                    }
                    data += "\r\n";

                }
            }
            try
            {
                File.WriteAllText(location, data);
                return objName.Length + " object saved";
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