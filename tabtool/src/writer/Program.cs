using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Saro.Table
{
    class Program
    {
        static async Task Main(string[] args)
        {
            await Load(args);
        }

        static async Task Load(string[] args)
        {
            int processorCount = System.Environment.ProcessorCount;
            ThreadPool.SetMinThreads(Math.Max(4, processorCount), 5);
            ThreadPool.SetMaxThreads(Math.Max(16, processorCount * 4), 10);

            string clientOutDir = null, serverOutDir = null, csOutDir = null, excelDir = null;
            CmdlineHelper cmder = new CmdlineHelper(args);
            if (cmder.Has("--out_client")) { clientOutDir = cmder.Get("--out_client"); } else { Console.WriteLine("out_client missing"); return; }
            ////if (cmder.Has("--out_server")) { serverOutDir = cmder.Get("--out_server"); } else { return; }
            if (cmder.Has("--in_excel")) { excelDir = cmder.Get("--in_excel"); } else { Console.WriteLine("in_excel missing"); return; }

            //Console.WriteLine(clientOutDir);
            //Console.WriteLine(excelDir);

            //创建导出目录
            if (!Directory.Exists(clientOutDir)) Directory.CreateDirectory(clientOutDir);
            //if (!Directory.Exists(serverOutDir)) Directory.CreateDirectory(serverOutDir);

            bool gen_client_cs = false;
            if (cmder.Has("--out_cs"))
            {
                csOutDir = cmder.Get("--out_cs");
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

                Console.WriteLine();
                Console.WriteLine("process xls: " + fileName);

#if true
                var excelDatas = TableHelper.LoadExcel_V1(filepath);

                tasks.Add(Task.Run(() =>
                {
                    foreach (var excelData in excelDatas)
                    {
                        string clientPath = clientOutDir + excelData.tablName + ".txt";
                        //string serverPath = serverOutDir + sheets[i].SheetName + ".txt";

                        Console.WriteLine("parsing...... " + excelData.tablName);

                        TableHelper.WriteByteAsset(excelData, clientPath);
                    }

                    if (gen_client_cs)
                    {
                        var codepath = csOutDir + "/" + fileName + ".cs";

                        CodeGen.MakeCsharpFile(excelDatas, codepath);
                    }
                }));
#endif

#if false
                var excelDatas = TableHelper.LoadExcelFileAsync_EPPlus(filepath);
                //tasks.Add(Task.Run(() =>
                //{
                //    foreach (var excelData in excelDatas)
                //    {
                //        Console.WriteLine();
                //        Console.WriteLine("parsing...... " + excelData.tablName);
                //        string clientPath = clientOutDir + excelData.tablName + ".txt";
                //        //string serverPath = serverOutDir + sheets[i].SheetName + ".txt";

                //        TableHelper.WriteByteAsset(excelData, clientPath);
                //    }


                //    if (cmder.Has("--out_cs"))
                //    {
                //        csOutDir = cmder.Get("--out_cs");
                //        if (!Directory.Exists(csOutDir))
                //            Directory.CreateDirectory(csOutDir);

                //        var codepath = csOutDir + "/" + fileName + ".cs";

                //        CodeGen.MakeCsharpFile(excelDatas, codepath);
                //        Console.WriteLine();
                //        Console.WriteLine($"generate {codepath} successful!");
                //    }
                //}));
#endif

#if false
                var sheets = TableHelper.LoadExcelFile(filepath);

                tasks.Add(Task.Run(() =>
                {
                    for (int i = 0; i < sheets.Count; i++)
                    {
                        Console.WriteLine();
                        Console.WriteLine("parsing...... " + sheets[i].SheetName);
                        string clientPath = clientOutDir + sheets[i].SheetName + ".txt";
                        //string serverPath = serverOutDir + sheets[i].SheetName + ".txt";

                        var data = TableHelper.ParseExcel(sheets[i]);
                        TableHelper.WriteByteAsset(data, clientPath);

                        if (cmder.Has("--out_cs"))
                        {
                            csOutDir = cmder.Get("--out_cs");
                            if (!Directory.Exists(csOutDir))
                                Directory.CreateDirectory(csOutDir);

                            var codepath = csOutDir + "/" + fileName + ".cs";

                            CodeGen.MakeCsharpFile(new List<ExcelData> { data }, codepath);
                            Console.WriteLine();
                            Console.WriteLine($"generate {codepath} successful!");
                        }
                    }
                }));
#endif
            }

            await Task.WhenAll(tasks);

            Console.WriteLine();
            Console.WriteLine("export success!");

            time.Stop();
            Console.WriteLine($"process finish: {time.ElapsedMilliseconds} ms");
        }
    }
}
