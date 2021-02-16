using System;
using System.IO;

namespace Saro.Table.sample
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

            csvTest.Get().Load();
            Console.WriteLine(csvTest.Get().PrintTable());

            csvTest1.Get().Load();
            Console.WriteLine(csvTest1.Get().PrintTable());

            Console.WriteLine(string.Join(",", csvTest2.Query(0, 0, 0).float_arr));

            csvTest2.Get().Unload();

            Console.ReadKey();
        }
    }
}
