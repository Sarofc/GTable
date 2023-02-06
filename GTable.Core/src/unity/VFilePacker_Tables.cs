#if UNITY_2017_1_OR_NEWER

using System;
using System.IO;
using Saro.IO;
using Saro.MoonAsset.Build;
using UnityEditor;

namespace Saro.GTable
{
    internal sealed class VFilePacker_Tables : IVFilePacker
    {
        bool IVFilePacker.PackVFile(string dstFolder)
        {
            var dstVFilePath = dstFolder + "/" + GTableConfig.vfileName;

            BuildVFile(dstVFilePath);

            return true;
        }

        private static void BuildVFile(string vfilePath)
        {
            try
            {
                if (File.Exists(vfilePath))
                    File.Delete(vfilePath);

                var tablePath = GTableConfig.out_client;
                if (!Directory.Exists(tablePath))
                {
                    Log.ERROR("没有数据表数据，跳过: " + tablePath);
                    return;
                }
                var files = Directory.GetFiles(tablePath, "*", SearchOption.AllDirectories);
                if (files.Length > 0)
                {
                    using (var vfile = VFileSystem.Open(vfilePath, FileMode.CreateNew, FileAccess.ReadWrite, files.Length, files.Length))
                    {
                        for (int i = 0; i < files.Length; i++)
                        {
                            string file = files[i];

                            EditorUtility.DisplayCancelableProgressBar("打包数据表", $"{file}", (i + 1f) / files.Length);

                            var result = vfile.WriteFile($"{Path.GetFileName(file)}", file);
                            if (!result)
                            {
                                Log.ERROR($"[VFilePacker_Tables] 打包数据表失败： {file} ");
                                continue;
                            }
                        }

                        Log.ERROR("[VFilePacker_Tables]\n" + string.Join("\n", vfile.GetAllFileInfos()));
                    }
                }
            }
            catch (Exception e)
            {
                Log.ERROR("[VFilePacker_Tables] 打包数据表失败. error:" + e);
                return;
            }
            finally
            {
                EditorUtility.ClearProgressBar();
            }
        }
    }
}

#endif