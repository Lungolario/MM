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
    }
    public class Model : DNAObject
    {
        public override void CreateObject(string name)
        {
            base.CreateObject(name);
        }
    }
    public class Curve : DNAObject
    {
        public override void CreateObject(string name)
        {
            base.CreateObject(name);
        }
    }
    public class Vol : DNAObject
    {
        public override void CreateObject(string name)
        {
            base.CreateObject(name);
        }
    }
    public class Results : DNAObject
    {
        public override void CreateObject(string name)
        {
            base.CreateObject(name);
        }
    }
    public class MMObjectHandler
    {
        public List<DNAObject> objs = new List<DNAObject>();
        public object[,] rangeForDisplay;
        public int CreateObject(string name, string type)
        {
            switch (type)
            {
                case "MODEL":
                    return CreateObject<Model>(name);
                case "CURVE":
                    return CreateObject<Curve>(name);
                case "VOL":
                    return CreateObject<Vol>(name);
                case "RESULTS":
                    return CreateObject<Results>(name);
                default:
                    throw new Exception("Cannot add this object. No type found!(" + type + ")");
            }
        }
        public int CreateObject<T>(string name) where T : DNAObject, new()
        {
            foreach (var obj2 in objs.OfType<T>())
            {
                if (obj2.name == name)
                {
                    obj2.counter++;
                    return obj2.counter;
                }
            }
            DNAObject obj = new T();
            obj.CreateObject(name);
            obj.counter++;
            objs.Add(obj);
            return obj.counter;
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
