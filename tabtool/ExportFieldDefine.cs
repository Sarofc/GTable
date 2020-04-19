using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tabtool
{
    internal class TabToolConfig
    {
        public static string[] InValidSheetNames = new string[]
        {
            "Sheet",
            "sheet"
        };

        public static Dictionary<string, int> HeaderDefine = new Dictionary<string, int>()
        {
            {"filter"  ,0},
            {"info"    ,1},
            {"type"    ,2},
            {"name"    ,3},
            {"value"   ,4},
        };

        internal class ExportFilter
        {
            public const string k_NOT = "not";
            public const string k_CLIENT = "client";
            public const string k_SERVER = "server";
        }
    }

}
