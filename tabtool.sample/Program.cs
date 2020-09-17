using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tabtool.sample
{
    class Program
    {
        public const string k_ConfigPath = @"..\..\generate\data\";

        static void Main(string[] args)
        {
            TableCfg.s_TableSrc = k_ConfigPath;

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

            CfgTest1.Get().Load();
            Console.WriteLine(CfgTest1.Get().ToString());

            Console.ReadKey();
        }
    }
}
