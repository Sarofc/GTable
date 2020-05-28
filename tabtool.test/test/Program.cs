//using NPOI.SS.UserModel;
//using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using tabtool;

namespace tabtool
{
    class Program
    {
        static void Main(string[] args)
        {
            TableCfg.s_TableSrc = @"O:\Git\Saro\tabtool\tabtool.test\config\";

            TableCfg.s_BytesLoader = path =>
            {
                using (var fs = new FileStream(path, FileMode.Open))
                {
                    var data = new byte[fs.Length];
                    fs.Read(data, 0, data.Length);
                    return data;
                }
            };

            CfgTest.Get().Load();

            Console.WriteLine(CfgTest.Get().ToString());
            Console.WriteLine(CfgTest.Get().GetTableItem((int)EIDTest.key1)._string);

            CfgTest.Get().Unload();

            Console.ReadKey();
        }

    }
}
