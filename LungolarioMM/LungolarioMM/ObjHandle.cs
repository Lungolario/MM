﻿using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDna.Integration;

namespace MMA
{
    public abstract class ObjHandle
    {
        private string _name;
        private int _counter;
        public void CreateObject(string name, object[,] range)
        {
            _name = name;
            if (range.GetLength(0) > 0 && range.GetLength(1) > 1)
            {
                FieldInfo[] keyList = GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
                for (int i = 0; i < range.GetLength(0); i++)
                {
                    if (range[i, 0] == ExcelEmpty.Value)
                        continue;
                    int j;
                    for (j = 0; j < keyList.Length; j++)
                        if (range[i, 0].ToString().ToUpper() == keyList[j].Name.ToUpper())
                        {
                            if (typeof(IIMatrix).IsAssignableFrom(keyList[j].FieldType))
                            {
                                int countRows, countColumns;
                                for (countRows = 1; countRows + i < range.GetLength(0); countRows++)
                                    if (range[countRows + i, 0] != ExcelEmpty.Value || range[countRows + i, 1] == ExcelEmpty.Value) break;
                                for (countColumns = 1; countColumns < range.GetLength(1); countColumns++)
                                    if (range[i + 1, countColumns] == ExcelEmpty.Value) break;

                                IIMatrix mat = (IIMatrix)Activator.CreateInstance(keyList[j].FieldType);
                                mat.CreateMatrix(range, i + 1, countRows - 1, 1, countColumns - 1);
                                keyList[j].SetValue(this, mat);
                                i += countRows - 1;
                            }
                            else if (range[i, 1] == ExcelEmpty.Value)
                                keyList[j].SetValue(this, null);
                            else
                            {
                                Type fieldType = keyList[j].FieldType;
                                if (Nullable.GetUnderlyingType(fieldType) != null)
                                    fieldType = Nullable.GetUnderlyingType(fieldType);
                                try
                                {
                                    if (fieldType.BaseType.Name == "Enum")
                                        keyList[j].SetValue(this, Enum.Parse(fieldType, range[i, 1].ToString(), true));
                                    else
                                        keyList[j].SetValue(this, Convert.ChangeType(range[i, 1], fieldType));
                                }
                                catch (Exception)
                                {
                                    throw new Exception(fieldType.Name + " cannot be set to " + range[i, 1]);
                                }
                            }
                            break;
                        }
                    if (j == keyList.Length)
                        throw new Exception("Key " + range[i, 0] + " not available for object " + GetType().Name);
                }
            }
            CheckObject();
        }

        protected virtual void CheckObject()
        {
        }

        public object[,] DisplayObject()
        {
            FieldInfo[] keyList = GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
            MatrixBuilder result = new MatrixBuilder();
            foreach (var key in keyList)
                if (key.GetValue(this) != null)
                {
                    result.Add(new[] { key.Name }, false, true, false);
                    if (typeof(IIMatrix).IsAssignableFrom(key.FieldType))
                        result.Add(((IIMatrix)key.GetValue(this)).ObjInfo(ExcelMissing.Value, ExcelMissing.Value), true, true, false);
                    else
                        result.Add(new[] { key.GetValue(this) }, true, false, false);
                }
            return result.Deliver();
        }
        public string Serialize()
        {
            object[,] objMatrix = DisplayObject();
            string data = "NEW" + GetType().Name.ToUpper() + "\r\n";
            data += "_name\t" + GetName() + "\r\n";
            for (int iRow = 0; iRow < objMatrix.GetLength(0); iRow++)
            {
                for (int iCol = 0; iCol < objMatrix.GetLength(1); iCol++)
                    data += objMatrix[iRow, iCol] + "\t";
                data += "\r\n";
            }
            return data;
        }

        public ObjHandle TakeOverOldObject(ObjectHandler objHandler)
        {
            ObjHandle existingObj = objHandler.GetObject(_name, GetType().Name);
            if (existingObj != null)
            {
                _counter = existingObj._counter + 1;
                objHandler.ObjList.Remove(existingObj);
            }
            return this;
        }

        public string ModifyObject(string key, object value)
        {
            FieldInfo[] keyList = GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
            foreach (var keyL in keyList)
                if (key.ToUpper() == keyL.Name.ToUpper())
                {
                    try
                    {
                        keyL.SetValue(this, value);
                        return FinishMod().GetNameCounter();
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }
                }
            return "Object found, key not found.";
        }

        public ObjHandle FinishMod()
        {
            _counter++;
            FieldInfo[] privateProps = GetType().GetFields(BindingFlags.NonPublic | BindingFlags.Instance);
            foreach (var privProp in privateProps)
                privProp.SetValue(this, null);
            return this;
        }
        public string GetName() => _name;
        public string GetNameCounter() => _name + ":" + _counter;
    }
    public class Model : ObjHandle
    {
        public string ModelName;
        public int ExtraResults;
    }
    public class Vol : ObjHandle
    {
        public string Currency;
        public double Volatility;
    }
    public class Results : ObjHandle
    {
        public double Result;
    }
    public class Trade : ObjHandle
    {
        public string Currency;
        public double Start;
    }
    public class Dictionary : ObjHandle
    {
        public string MODE;
        public string MODEL;
        public string CURVE;
        public string VOL;
        public string TRADE;
        public string RESULTS;
    }
    public class ObjectHandler
    {
        public List<ObjHandle> ObjList = new List<ObjHandle>();
        public ObjHandle CreateOrOverwriteObject(string name, string type, object[,] range)
        {
            ObjHandle newObj = (ObjHandle)Activator.CreateInstance(Type.GetType(typeof(ObjHandle).Namespace + "." + type, true, true));
            newObj.CreateObject(Tools.StringTrim(name), range);
            ObjList.Add(newObj.TakeOverOldObject(this));
            return newObj;
        }
        public ObjHandle GetObject(string name, string type)
        {
            return ObjList.Find(item => item.GetName().ToUpper().Equals(Tools.StringTrim(name).ToUpper()) && item.GetType().Name.ToUpper() == type.ToUpper());
        }
        public int DeleteObjects(string name, string type)
        {
            int i = ObjList.Count;
            if (type.Length > 0  && name.Length > 0)
                ObjList.RemoveAll(item => item.GetName().ToUpper().Equals(name.ToUpper()) && item.GetType().Name.ToUpper() == type.ToUpper());
            else if (type.Length > 0)
                ObjList.RemoveAll(item => item.GetType().Name.ToUpper() == type.ToUpper());
            else
                ObjList.Clear();
            return i - ObjList.Count;
        }
    }
}