using BenchmarkDotNet.Running;
using System;
using System.Collections.Generic;
using System.IO;

namespace Saro.Table.sample
{
    class Program
    {
        public const string k_ConfigPath = @"..\..\..\generate\data\";

        static void Main(string[] args)
        {
            TestTable();

            //BenchmarkRunner.Run<Bench_Split>();
        }
        

        private static void TestTable()
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

            csvTest1.Get().Load();
            Console.WriteLine(csvTest1.Get().PrintTable());

            Console.WriteLine(string.Join(",", csvTest1.Query(0, 0, 0).float_arr));
            Console.WriteLine(string.Join(",", csvTest1.Query(0, 0, 0).map_int_int));

            csvTest1.Get().Unload();

            Console.ReadKey();
        }
    }
}
