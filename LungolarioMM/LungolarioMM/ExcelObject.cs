using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDna.Integration;

namespace MMA
{
    public abstract class ExcelObject
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
                                    if (range[countRows + i, 0].GetType() != typeof(ExcelEmpty) || range[countRows + i, 1].GetType() == typeof(ExcelEmpty)) break;
                                for (countColumns = 1; countColumns < range.GetLength(1); countColumns++)
                                    if (range[i + 1, countColumns].GetType() == typeof(ExcelEmpty)) break;

                                IIMatrix mat = (IIMatrix)Activator.CreateInstance(keyList[j].FieldType);
                                mat.CreateMatrix(range, i + 1, countRows - 1, 1, countColumns - 1);
                                keyList[j].SetValue(this, mat);
                                i += countRows - 1;
                            }
                            else if (range[i, 1] == ExcelEmpty.Value)
                                keyList[j].SetValue(this, null);
                            else if (Nullable.GetUnderlyingType(keyList[j].FieldType) == null)
                            {
                                var val = Convert.ChangeType(range[i, 1], keyList[j].FieldType);
                                keyList[j].SetValue(this, val);
                            }
                            else
                            {
                                var val = Convert.ChangeType(range[i, 1], Nullable.GetUnderlyingType(keyList[j].FieldType));
                                keyList[j].SetValue(this, val);
                            }
                            break;
                        }
                    if (j == keyList.Length)
                        throw new Exception("Key " + range[i, 0] + " not available for object " + GetType().Name);
                }
            }
            CheckObject();
        }

        public virtual void CheckObject()
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
        public string ObjectSerialize()
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

        public ExcelObject TakeOverOldObject(ExcelObjectHandler objHandler)
        {
            ExcelObject existingObj = objHandler.GetObject(_name, GetType().Name);
            if (existingObj != null)
            {
                _counter = existingObj._counter + 1;
                objHandler.ObjList.Remove(existingObj);
            }
            return this;
        }
        public ExcelObject FinishMod()
        {
            _counter++;
            FieldInfo[] privateProps = GetType().GetFields(BindingFlags.NonPublic | BindingFlags.Instance);
            foreach (var privProp in privateProps)
                privProp.SetValue(this, null);
            return this;
        }
        public string GetName() { return _name; }
        public string GetNameCounter() { return _name + ":" + _counter; }
    }
    public class Model : ExcelObject
    {
        public string modelName;
        public int extraResults;
    }
    public class Vol : ExcelObject
    {
        public string currency;
        public double volatility;
    }
    public class Results : ExcelObject
    {
        public double result;
    }
    public class Security : ExcelObject
    {
        public string currency;
        public double start;
    }
    public class Dictionary : ExcelObject
    {
        public string MODE;
        public string MODEL;
        public string CURVE;
        public string VOL;
        public string SECURITY;
        public string RESULTS;
    }
    public class ExcelObjectHandler
    {
        public List<ExcelObject> ObjList = new List<ExcelObject>();
        public ExcelObject CreateObject(string name, string type, object[,] range)
        {
            ExcelObject newObj = (ExcelObject)Activator.CreateInstance(Type.GetType(typeof(ExcelObject).Namespace + "." + type, true, true));
            newObj.CreateObject(Tools.StringTrim(name), range);
            ObjList.Add(newObj.TakeOverOldObject(this));
            return newObj;
        }
        public ExcelObject GetObject(string name, string type)
        {
            return ObjList.Find(item => item.GetName().ToUpper().Equals(Tools.StringTrim(name).ToUpper()) && item.GetType().Name.ToUpper() == type.ToUpper());
        }
    }
}