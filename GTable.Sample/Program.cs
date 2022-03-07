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
            //TestTable();

            var key0 = KeyHelper.GetKey(-10, -20);
            Console.WriteLine($"{key0} {KeyHelper.SplitKey2(key0)}");

            var key1 = KeyHelper.GetKey(-10, -20, -30);
            Console.WriteLine($"{key1} {KeyHelper.SplitKey3(key1)}");

            var key2 = KeyHelper.GetKey(-10, -20, -30, -40);
            Console.WriteLine($"{key2} {KeyHelper.SplitKey4(key2)}");


            var key3 = KeyCombine(-10, -20);
            var result = SplitKey(key3);

            Console.WriteLine($"{key3} {result}");

            //BenchmarkRunner.Run<Bench_Split>();
        }

        private static ulong KeyCombine(int key1, int key2)
        {
            // Note: if you're in a checked context by default, you'll want to make this
            // explicitly unchecked
            uint u1 = (uint)key1;
            uint u2 = (uint)key2;

            ulong unsignedKey = u1 | (((ulong)u2) << 32);
            return unsignedKey;
        }

        private static (int key1, int key2) SplitKey(ulong key)
        {
            //And to reverse:
            ulong unsignedKey = key;
            uint highBits = (uint)(unsignedKey & 0xffffffffUL);
            uint lowBits = (uint)(unsignedKey >> 32);
            int i1 = (int)highBits;
            int i2 = (int)lowBits;

            return (i1, i2);
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
