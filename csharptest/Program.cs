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
        static void Main(string[] args)
        {
            //var path = @"O:\Git\Saro\tabtool\csharptest\config\achive.txt";

            //FileStream fs = new FileStream(path, FileMode.Open);
            //BinaryFormatter bf = new BinaryFormatter();
            //DataTable dt = bf.Deserialize(fs) as DataTable;

            //for (int i1 = 0; i1 < dt.Rows.Count; i1++)
            //{
            //    if (i1 == 0 || i1 == 1 || /*i == 2 ||*/ i1 == 3) continue;
            //    for (int j = 1; j < dt.Columns.Count; j++)
            //    {
            //        if (j == dt.Columns.Count - 1)
            //        {
            //            Console.Write(dt.Rows[i1].ItemArray[j].ToString());
            //        }
            //        else
            //        {
            //            Console.Write(dt.Rows[i1].ItemArray[j].ToString() + "\t");
            //        }
            //    }
            //    Console.WriteLine();
            //}
            var time = new Stopwatch();
            time.Start();
            if (CfgachiveTable.Instance.Load())
            {
                TbsachiveItem item = CfgachiveTable.Instance.GetTableItem(1);
                if (item != null)
                {
                    Console.WriteLine(item);
                }
            }
            time.Stop();
            Console.WriteLine("Load huge table: " + time.ElapsedMilliseconds / 1000f);

            Console.ReadKey();
        }
    }
}
