using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Xml;
using System.IO;

namespace tabtool
{
    class Program
    {
        static void Main(string[] args)
        {
            Load(args);

            Console.ReadKey();
        }

        static void Load(string[] args)
        {
            string clientOutDir, serverOutDir, csOutDir, excelDir, metafile;
            CmdlineHelper cmder = new CmdlineHelper(args);
            if (cmder.Has("--out_client")) { clientOutDir = cmder.Get("--out_client"); } else { return; }
            if (cmder.Has("--out_server")) { serverOutDir = cmder.Get("--out_server"); } else { return; }
            if (cmder.Has("--in_excel")) { excelDir = cmder.Get("--in_excel"); } else { return; }
            if (cmder.Has("--in_tbs")) { metafile = cmder.Get("--in_tbs"); } else { return; }

            //创建导出目录
            if (!Directory.Exists(clientOutDir)) Directory.CreateDirectory(clientOutDir);
            if (!Directory.Exists(serverOutDir)) Directory.CreateDirectory(serverOutDir);

            //先读取tablemata文件
            TableStruct tbs = new TableStruct();
            if (!tbs.ImportTableStruct(metafile))
            {
                Console.WriteLine("解析tbs文件错误！");
                return;
            }
            Console.WriteLine("解析tbs文件成功");

            List<TableMeta> clientTableMetaList = new List<TableMeta>();

            //导出文件
            TableHelper helper = new TableHelper();
            string[] files = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly);
            foreach (string filepath in files)
            {
                try
                {
                    var sheets = helper.ImportExcelFile(filepath);

                    for (int i = 0; i < sheets.Count; i++)
                    {
                        string clientPath = clientOutDir + sheets[i].SheetName + ".txt";
                        string serverPath = serverOutDir + sheets[i].SheetName + ".txt";

                        var sb = new StringBuilder();
                        var dt = helper.GetDataTable(sheets[0]);
                        helper.WriteTxtAsset(dt, clientPath);
                        //helper.WriteByteAsset(dt, clientPath);
                        var meta = helper.GetTableMeta(clientPath, dt);
                        clientTableMetaList.Add(meta);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(filepath + " 文件出错！");
                    Console.WriteLine(e.ToString());
                }
            }
            Console.WriteLine("导出配置文件成功");

            if (cmder.Has("--out_cs"))
            {
                csOutDir = cmder.Get("--out_cs");
                if (!Directory.Exists(csOutDir))
                    Directory.CreateDirectory(csOutDir);

                CodeGen.MakeCsharpFileTbs(tbs.GetMetaList(), csOutDir);
                CodeGen.MakeCsharpFile(clientTableMetaList, csOutDir);
                Console.WriteLine("生成.cs代码文件成功");
            }

            Console.WriteLine("按任意键退出...");
            Console.ReadKey(false);
        }
    }
}
