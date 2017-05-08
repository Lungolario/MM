using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace LungolarioMM
{
    public class FirstAddIn
    {

        [ExcelFunction(Description = "Adds Hello World to the start of the string given as input")]
        public static string HelloWorld(string name)
        {
            return "Hello World " + name;
        }

    }
}
