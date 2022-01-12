using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace Saro.Table.sample
{
    class Program
    {
        public const string k_ConfigPath = @"..\..\..\generate\data\";

        static void Main(string[] args)
        {
            //Test();

            RenameSheets();
        }

        private static void RenameSheets()
        {
            var excels = @"C:\Users\sarof\Projects\Git\tabtool\tables\excel";

            string[] files = Directory.GetFiles(excels, "*.xlsx", SearchOption.TopDirectoryOnly);

            var sheetIndex = 0;

            foreach (string filepath in files)
            {
                IWorkbook workbook;

                using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    workbook = new XSSFWorkbook(fs);
                }

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);

                    workbook.SetSheetName(i, "Test" + sheetIndex++);

                    Console.WriteLine($"rename: {sheet.SheetName}");
                }

                using (var fs1 = new FileStream(filepath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                {
                    workbook.Write(fs1);
                }
            }
        }

        private static void Test()
        {
            //TableCfg.s_TableSrc = k_ConfigPath;

            //TableCfg.s_BytesLoader = path =>
            //{
            //    using (var fs = new FileStream(path, FileMode.Open))
            //    {
            //        var data = new byte[fs.Length];
            //        fs.Read(data, 0, data.Length);
            //        return data;
            //    }
            //};

            //csvTest.Get().Load();
            //Console.WriteLine(csvTest.Get().PrintTable());

            //csvTest1.Get().Load();
            //Console.WriteLine(csvTest1.Get().PrintTable());

            //Console.WriteLine(string.Join(",", csvTest2.Query(0, 0, 0).float_arr));

            //csvTest2.Get().Unload();

            Console.ReadKey();
        }
    }
}
