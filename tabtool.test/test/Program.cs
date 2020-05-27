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
        public enum ETableFieldType : byte
        {
            Byte,
            Int,
            Long,
            Float,
            String,
            Struct,// TODO
            ByteList,
            IntList,
            LongList,
            FloatList,
            StructList,// TODO
        }

        class Item
        {
            public int id;
            public string value;
        }

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

            CfgITEMTable.Get().Load();

            Console.WriteLine(CfgITEMTable.Get().ToString());
            Console.WriteLine(CfgITEMTable.Get().GetTableItem(1).name);

            CfgITEMTable.Get().Unload();

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
