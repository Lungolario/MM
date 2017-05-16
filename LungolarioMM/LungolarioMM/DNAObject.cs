using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LungolarioMM
{
    public abstract class DNAObject
    {
        public string name;
        public int counter;
        public virtual void CreateObject(string name)
        {
            this.name = name;
        }
        public abstract string ReturnType();
    }
    public class Model : DNAObject
    {
        public override void CreateObject(string name)
        {
            base.CreateObject(name);
        }
        public override string ReturnType()
        {
            return "Model";
        }
    }
    public class Curve : DNAObject
    {
        public override void CreateObject(string name)
        {
            base.CreateObject(name);
        }
        public override string ReturnType()
        {
            return "Curve";
        }
    }
    public class Vol : DNAObject
    {
        public override void CreateObject(string name)
        {
            base.CreateObject(name);
        }
        public override string ReturnType()
        {
            return "Vol";
        }
    }
    public class Results : DNAObject
    {
        public override void CreateObject(string name)
        {
            base.CreateObject(name);
        }
        public override string ReturnType()
        {
            return "Results";
        }
    }
    public class MMObjectHandler
    {
        public List<DNAObject> objs = new List<DNAObject>();
        public int startRow,startCol,endRow,endCol;
        public string startRef, endRef;
        public void CreateObject(string name, string type)
        {
            bool found = false;
            switch (type)
            {
                case "MODEL":
                    foreach (var obj in objs.OfType<Model>())
                    {
                        if (obj.name == name)
                        {
                            obj.counter++;
                            found = true;
                        }
                    }
                    if (!found)
                    {
                        DNAObject obj = new Model();
                        obj.CreateObject(name);
                        obj.counter++;
                        objs.Add(obj);
                    }
                    break;
                case "CURVE":
                    foreach (var obj in objs.OfType<Curve>())
                    {
                        if (obj.name == name)
                        {
                            obj.counter++;
                            found = true;
                        }
                    }
                    if (!found)
                    {
                        DNAObject obj = new Curve();
                        obj.CreateObject(name);
                        obj.counter++;
                        objs.Add(obj);
                    }
                    break;
                case "VOL":
                    foreach (var obj in objs.OfType<Vol>())
                    {
                        if (obj.name == name)
                        {
                            obj.counter++;
                            found = true;
                        }
                    }
                    if (!found)
                    {
                        DNAObject obj = new Vol();
                        obj.CreateObject(name);
                        obj.counter++;
                        objs.Add(obj);
                    }
                    break;
                case "RESULTS":
                    foreach (var obj in objs.OfType<Results>())
                    {
                        if (obj.name == name)
                        {
                            obj.counter++;
                            found = true;
                        }
                    }
                    if (!found)
                    {
                        DNAObject obj = new Results();
                        obj.CreateObject(name);
                        obj.counter++;
                        objs.Add(obj);
                    }
                    break;
                default:
                    throw new Exception("Cannot add this object. No type found!(" + type + ")");
            }

        }
        public string CountAllObjects()
        {
            string countedObjs = "";

            foreach (var x in objs)
            {
                if (x.counter > 0)
                    countedObjs += x.name + ":" + x.counter + " ";
            }
            return countedObjs;
        }
    }
}








/*****************
 * Don't delete this code. Can to be used later on if something is wrong.
***************************************************/


//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;

//namespace AbstractClassTest
//{
//    public abstract class DNAObject
//    {
//        public string name;
//        public virtual void CreateObject(string name)
//        {
//            this.name = name;
//        }
//    }
//    public class Model : DNAObject
//    {
//        public override void CreateObject(string name)
//        {
//            base.CreateObject(name);
//        }
//    }
//    public class Curve : DNAObject
//    {
//        public override void CreateObject(string name)
//        {
//            base.CreateObject(name);
//        }
//    }
//    public class Vol : DNAObject
//    {
//        public override void CreateObject(string name)
//        {
//            base.CreateObject(name);
//        }
//    }
//    public class Results : DNAObject
//    {
//        public override void CreateObject(string name)
//        {
//            base.CreateObject(name);
//        }
//    }
//    public class MMObjectHandler
//    {
//        public List<DNAObject> objs = new List<DNAObject>();
//        public void CreateObject(string name, string type)
//        {
//            switch (type)
//            {
//                case "MODEL":
//                    DNAObject model = new Model();
//                    model.name = name;
//                    objs.Add(model);
//                    break;
//                case "CURVE":
//                    DNAObject curve = new Curve();
//                    curve.name = name;
//                    objs.Add(curve);
//                    break;
//                case "VOL":
//                    DNAObject vol = new Vol();
//                    vol.name = name;
//                    objs.Add(vol);
//                    break;
//                case "RESULTS":
//                    DNAObject results = new Results();
//                    results.name = name;
//                    objs.Add(results);
//                    break;
//                default:
//                    throw new Exception("No type found!");
//            }

//        }
//        public int CountObjectByName(string name)
//        {
//            return objs.Count(x => x.name == name);
//        }
//        public string CountAllObjects()
//        {
//            string countedObjs = "";
//            var q = from x in objs
//                    group x by x.name into g
//                    let count = g.Count()
//                    orderby count descending
//                    select new { Name = g.Key, Count = count };
//            foreach (var x in q)
//            {
//                countedObjs += x.Name + ":" + x.Count;
//            }
//            return countedObjs;
//        }
//    }
//}
