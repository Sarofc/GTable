using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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

namespace csharptest
{
    class Program
    {
        class Item
        {
            public int id;
            public string value;
            public List<int> test;
        }

        static void Main(string[] args)
        {
            var path = @"O:\Git\Saro\MGFTemplate\tabtool\data\config\EN.txt";

            int version;
            byte[] types;

            var list = new List<Item>();

            using (var fs = new FileStream(path, FileMode.Open))
            {
                var br = new BinaryReader(fs, Encoding.UTF8);
                version = br.ReadInt32();//version
                var typeCount = br.ReadInt32();
                types = new byte[typeCount];
                for (int i = 0; i < typeCount; i++)
                {
                    types[i] = br.ReadByte();
                }

                var dataLen = br.ReadInt32();

                for (int i = 0; i < dataLen; i++)
                {
                    var item = new Item();
                    item.id = br.ReadInt32();
                    item.value = br.ReadString();

                    var testCount = br.ReadInt32();
                    item.test = new List<int>(testCount);
                    for (int i1 = 0; i1 < testCount; i1++)
                    {
                        item.test.Add(br.ReadInt32());
                    }

                    list.Add(item);
                }
            }

            var sb = new StringBuilder();
            sb.Append(version).AppendLine();
            foreach (var type in types)
            {
                sb.Append(type).Append("\t");
            }
            sb.AppendLine();
            foreach (var item in list)
            {
                sb.Append(item.id).Append(": ").Append(item.value).
                    Append(" | ").
                    Append(string.Join(",", item.test)).
                    AppendLine();
            }

            Console.WriteLine(sb.ToString());

            Console.ReadKey();
        }

        static void Test()
        {
            //var time = new Stopwatch();
            //time.Start();
            //if (CfgachiveTable.Instance.Load())
            //{
            //    TbsachiveItem item = CfgachiveTable.Instance.GetTableItem(1);
            //    if (item != null)
            //    {
            //        Console.WriteLine(item);
            //    }
            //}
            //time.Stop();
            //Console.WriteLine("Load huge table: " + time.ElapsedMilliseconds / 1000f);

            //Console.ReadKey();
        }
    }
}
