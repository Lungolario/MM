using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LungolarioMM
{
    public abstract class DNAObject
    {
        public string name;
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
        public void CreateObject(string name,string type)
        {
            switch(type)
            {
                case "MODEL":
                    DNAObject model = new Model();
                    model.name = name;
                    objs.Add(model);
                    break;
                case "CURVE":
                    DNAObject curve = new Curve();
                    curve.name = name;
                    objs.Add(curve);
                    break;
                case "VOL":
                    DNAObject vol = new Vol();
                    vol.name = name;
                    objs.Add(vol);
                    break;
                case "RESULTS":
                    DNAObject results = new Results();
                    results.name = name;
                    objs.Add(results);
                    break;
                default:
                    throw new Exception("Cannot add this object. No type found!(" + type +")");
            }
            
        }
        public int CountObjectByName(string name)
        {
            return objs.Count(x => x.name == name);
        }
        public string CountAllObjects()
        {
            string countedObjs="";
            var q = from x in objs
                    group x by x.name into g
                    let count = g.Count()
                    orderby count descending
                    select new { Value = g.Key, Count = g.Count() };
            foreach (var x in q)
            {
                countedObjs += x.Value + ":" + x.Count +" ";
            }
            return countedObjs;
        }
    }
}
