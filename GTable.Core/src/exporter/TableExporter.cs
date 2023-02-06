using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Saro.GTable
{
    public class TableExporter
    {
        /// <summary>
        /// --out_client [TABLE_DATA] --out_cs [TABLE_CS] --in_excel [TABLE_EXCEL]
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public static async Task ExportAsync(params string[] args)
        {
            int processorCount = Environment.ProcessorCount;
            ThreadPool.SetMinThreads(Math.Max(4, processorCount), 5);
            ThreadPool.SetMaxThreads(Math.Max(16, processorCount * 4), 10);

            string clientOutDir = null, serverOutDir = null, csOutDir = null, excelDir = null;
            CmdlineHelper cmder = new CmdlineHelper(args);
            if (cmder.Has("--out_client")) { clientOutDir = cmder.Get("--out_client"); } else { Log.Info("out_client missing"); return; }
            ////if (cmder.Has("--out_server")) { serverOutDir = cmder.Get("--out_server"); } else { return; }
            if (cmder.Has("--in_excel")) { excelDir = cmder.Get("--in_excel"); } else { Log.Info("in_excel missing"); return; }

            //Log.LogInfo(clientOutDir);
            //Log.LogInfo(excelDir);

            if (!Directory.Exists(excelDir))
            {
                Log.Info($"{excelDir} not found");
                return;
            }

            //创建导出目录
            if (!Directory.Exists(clientOutDir)) Directory.CreateDirectory(clientOutDir);
            //if (!Directory.Exists(serverOutDir)) Directory.CreateDirectory(serverOutDir);

            bool gen_client_cs = false;
            if (cmder.Has("--out_cs"))
            {
                csOutDir = cmder.Get("--out_cs");

                // TODO 这个操作很危险，最好能移动到回收站
                // if(Directory.Exists(csOutDir))
                //    Directory.Delete(csOutDir, true);

                if (!Directory.Exists(csOutDir))
                    Directory.CreateDirectory(csOutDir);

                gen_client_cs = true;
            }

            string[] files = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly);
            var tasks = new List<Task>(files.Length * 4);
            var time = new Stopwatch();
            time.Start();
            foreach (string filepath in files)
            {
                var fileName = Path.GetFileName(filepath);
                if (fileName.StartsWith("~")) continue;

                //Log.LogInfo($"[program] load excel {filepath}");

                var excelDatas = TableHelper.LoadExcel(filepath);

                tasks.Add(Task.Run(() =>
                {
                    foreach (var excelData in excelDatas)
                    {
                        string clientPath = clientOutDir + excelData.tablName + ".txt";
                        //string serverPath = serverOutDir + sheets[i].SheetName + ".txt";

                        TableHelper.WriteByteAsset(excelData, clientPath);

                        if (gen_client_cs)
                        {
                            var codepath = csOutDir + "/t" + excelData.tablName + ".g.cs";
                            CodeGen.MakeCsharpFile(excelData, codepath);
                        }
                    }
                }));
            }

            await Task.WhenAll(tasks);

            Log.Info("");
            Log.Info("export success!");

            time.Stop();
            Log.Info($"process finish: {time.ElapsedMilliseconds} ms");
        }
    }
}
