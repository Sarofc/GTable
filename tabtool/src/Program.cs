using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Xml;
using System.IO;
using System.Diagnostics;

namespace tabtool
{
    class Program
    {
        static void Main(string[] args)
        {
            Load(args);

        }

        static void Load(string[] args)
        {
            string clientOutDir, serverOutDir, csOutDir, excelDir, metafile;
            CmdlineHelper cmder = new CmdlineHelper(args);
            //if (cmder.Has("--out_client")) { clientOutDir = cmder.Get("--out_client"); } else { Console.WriteLine("out_client missing"); return; }
            ////if (cmder.Has("--out_server")) { serverOutDir = cmder.Get("--out_server"); } else { return; }
            if (cmder.Has("--in_excel")) { excelDir = cmder.Get("--in_excel"); } else { Console.WriteLine("in_excel missing"); return; }
            //if (cmder.Has("--in_tbs")) { metafile = cmder.Get("--in_tbs"); } else { Console.WriteLine("in_tbs missing"); return; }

            //Console.WriteLine(clientOutDir);
            //Console.WriteLine(excelDir);
            //Console.WriteLine(metafile);

            //创建导出目录
            //if (!Directory.Exists(clientOutDir)) Directory.CreateDirectory(clientOutDir);
            //if (!Directory.Exists(serverOutDir)) Directory.CreateDirectory(serverOutDir);

            //先读取tablemata文件
            //TableStruct tbs = new TableStruct();
            //if (!tbs.ImportTableStruct(metafile))
            //{
            //    Console.WriteLine("parse tbs file error！");
            //    return;
            //}
            //Console.WriteLine();
            //Console.WriteLine("parse tbs file successful!");

            List<ExcelData> clientExcelDataList = new List<ExcelData>();

            //导出文件
            TableHelper helper = new TableHelper();

            var sb = new StringBuilder();
            for (int i = 0; i < 100; i++)
            {
                if (i % 26 == 0) sb.AppendLine();
                sb.Append(helper.ConvertIntToLetter(i)).Append("\t");
            }
            Console.WriteLine(sb.ToString());

            string[] files = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly);
            //var time = new Stopwatch();
            //time.Start();
            foreach (string filepath in files)
            {
                var fileName = Path.GetFileName(filepath);
                if (fileName.StartsWith("~")) continue;

                Console.WriteLine();
                Console.WriteLine("process xls: " + fileName);

                try
                {
                    var sheets = helper.LoadExcelFile(filepath);

                    for (int i = 0; i < sheets.Count; i++)
                    {
                        //string clientPath = clientOutDir + sheets[i].SheetName + ".txt";
                        //string serverPath = serverOutDir + sheets[i].SheetName + ".txt";

                        var data = helper.ParseExcel(sheets[i]);

                        Console.WriteLine(data.ToString());

                        //helper.WriteTxtAsset(data, clientPath);
                        //helper.WriteByteAsset(data, clientPath);

                        clientExcelDataList.Add(data);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("export error！" + filepath);
                    Console.WriteLine(e.ToString());
                }
            }
            //time.Stop();
            //Console.WriteLine("end: " + time.ElapsedMilliseconds);
            Console.WriteLine("export success!");

            //if (cmder.Has("--out_cs"))
            //{
            //    csOutDir = cmder.Get("--out_cs");
            //    if (!Directory.Exists(csOutDir))
            //        Directory.CreateDirectory(csOutDir);

            //    //CodeGen.MakeCsharpFileTbs(tbs.GetMetaList(), csOutDir);
            //    CodeGen.MakeCsharpFile(clientExcelDataList, csOutDir);
            //    Console.WriteLine();
            //    Console.WriteLine("generate .cs code successful!");
            //}
        }
    }
}
