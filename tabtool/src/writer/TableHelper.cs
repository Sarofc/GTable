using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Formula.Functions;

namespace tabtool
{
    public class TableHelper
    {
        /// <summary>
        /// 解析Excel
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        internal ExcelData ParseExcel(ISheet sheet)
        {
            var defineRow = sheet.GetRow(0) as XSSFRow;
            var commentRow = sheet.GetRow(1) as XSSFRow;
            var typeRow = sheet.GetRow(2) as XSSFRow;
            var nameRow = sheet.GetRow(3) as XSSFRow;

            var data = new ExcelData();
            data.tablName = sheet.SheetName;
            var colCount = nameRow.LastCellNum - 1;

            data.header = new List<ExcelData.Header>(colCount);
            for (int i = 0; i < colCount; i++)
            {
                data.header.Add(new ExcelData.Header());
            }

            // 以name为基准
            int cellNum = nameRow.LastCellNum;
            int index = 1;
            for (; index < nameRow.LastCellNum; index++)
            {
                var cell = nameRow.GetCell(index);
                if (cell != null && !string.IsNullOrEmpty(cell.ToString()))
                {
                    data.header[index - 1].fieldName = (cell.ToString());
                }
                else
                {
                    Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{4}{ConvertIntToOrderedLetter(index)}]");
                    break;
                }

                cell = defineRow.GetCell(index);
                if (cell != null)
                    data.header[index - 1].define = (cell.ToString());
                else
                    Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{1}{ConvertIntToOrderedLetter(index)}]");

                cell = commentRow.GetCell(index);
                if (cell != null)
                    data.header[index - 1].fieldComment = (cell.ToString());
                else
                    Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{2}{ConvertIntToOrderedLetter(index)}]");

                cell = typeRow.GetCell(index);
                if (cell != null)
                {
                    data.header[index - 1].fieldTypeName = (cell.ToString());
                }
                else
                    Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{3}{ConvertIntToOrderedLetter(index)}]");
            }

            cellNum = index - 1;

            while (data.header.Count > cellNum)
            {
                data.header.RemoveAt(data.header.Count - 1);
            }

            //Console.WriteLine("CellNum: " + cellNum + "  NameRowNum: " + nameRow.LastCellNum);

            data.rowValues = new List<List<string>>(cellNum);
            for (int i = 4; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i) as XSSFRow;
                var keyCell = row.GetCell(1);
                if ((keyCell == null || string.IsNullOrEmpty(keyCell.ToString())))
                {
                    Console.WriteLine($"key [{i + 1}B] is null or empty. break.");
                    break;
                }

                data.rowValues.Add(new List<string>(cellNum));

                for (int j = 1; j <= cellNum; j++)
                {
                    var cell = row.GetCell(j);

                    if (cell == null)
                    {
                        data.rowValues[i - 4].Add(string.Empty);
                        Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{i + 1}{ConvertIntToOrderedLetter(i)}]");

                    }
                    else
                    {
                        data.rowValues[i - 4].Add(cell.ToString());
                    }
                }
            }

            //Console.WriteLine(data.ToString());

            return data;
        }

        /// <summary>
        /// 加载Excel表里的所有可用Sheet
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        internal List<ISheet> LoadExcelFile(string filePath)
        {
            IWorkbook m_Hssfworkbook;

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    m_Hssfworkbook = new XSSFWorkbook(fs);
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            List<ISheet> sheets = new List<ISheet>();

            for (int i = 0; i < m_Hssfworkbook.NumberOfSheets; i++)
            {
                var sheet = m_Hssfworkbook.GetSheetAt(i);
                if (!IsSheetNameValid(sheet))
                {
                    Console.WriteLine($"{Path.GetFileNameWithoutExtension(filePath)} skip sheet: {sheet.SheetName}");
                    continue;
                }
                sheets.Add(sheet);
            }

            return sheets;
        }

        /// <summary>
        /// 判断Sheet命名是否合法
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        internal bool IsSheetNameValid(ISheet sheet)
        {
            if (sheet != null)
            {
                foreach (var name in InValidSheetNames)
                {
                    if (sheet.SheetName.Contains(name))
                    {
                        return false;
                    }
                }

                if (sheet.SheetName.StartsWith("~"))
                {
                    return false;
                }

                return true;
            }

            return false;
        }

        [Obsolete("", true)]
        internal void WriteTxtAsset(ExcelData data, string path)
        {

        }

        /// <summary>
        /// 写入bytes
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        internal void WriteByteAsset(in ExcelData data, string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                BinaryWriter bw = new BinaryWriter(fs, Encoding.UTF8);

                bw.Write(ExcelData.k_DataVersion);
                bw.Write(data.rowValues.Count);//行数据长度

                for (int i = 0; i < data.rowValues.Count; i++)
                {
                    var line = data.rowValues[i];
                    for (int j = 0; j < data.header.Count; j++)
                    {
                        if (IgnoreHeader(data.header[j])) continue;

                        if (!s_TypeLut.TryGetValue(data.header[j].fieldTypeName, out Type t))
                        {
                            throw new Exception("type is not support: " + data.header[j].fieldTypeName);
                        }

                        if (t == typeof(byte))
                        {
                            var res = byte.TryParse(line[j], out byte val);
                            if (res) bw.Write(val);
                        }
                        else if (t == typeof(int))
                        {
                            var res = int.TryParse(line[j], out int val);
                            if (res) bw.Write(val);
                        }
                        else if (t == typeof(long))
                        {
                            var res = long.TryParse(line[j], out long val);
                            if (res) bw.Write(val);
                        }
                        else if (t == typeof(float))
                        {
                            var res = float.TryParse(line[j], out float val);
                            if (res) bw.Write(val);
                        }
                        else if (t == typeof(string))
                        {
                            bw.Write(line[j]);
                        }
                        else if (t == typeof(List<byte>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var bvs = line[j].Split(',');
                            bw.Write((ushort)bvs.Length);//长度
                            for (ushort i1 = 0; i1 < bvs.Length; i1++)
                            {
                                var res = byte.TryParse(bvs[i1].Trim(), out byte val);
                                if (res) bw.Write(val);
                            }
                        }
                        else if (t == typeof(List<int>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var ivs = line[j].Split(',');
                            bw.Write((ushort)ivs.Length);//长度
                            for (ushort i1 = 0; i1 < ivs.Length; i1++)
                            {
                                var res = int.TryParse(ivs[i1].Trim(), out int val);
                                if (res) bw.Write(val);
                            }
                        }
                        else if (t == typeof(List<long>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var lvs = line[j].Split(',');
                            bw.Write((ushort)lvs.Length);//长度
                            for (ushort i1 = 0; i1 < lvs.Length; i1++)
                            {
                                var res = long.TryParse(lvs[i1].Trim(), out long val);
                                if (res) bw.Write(val);
                            }
                        }
                        else if (t == typeof(List<float>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var fvs = line[j].Split(',');
                            bw.Write((ushort)fvs.Length);//长度
                            for (ushort i1 = 0; i1 < fvs.Length; i1++)
                            {
                                var res = float.TryParse(fvs[i1].Trim(), out float val);
                                if (res) bw.Write(val);
                            }
                        }
                        else if (t == typeof(Dictionary<int, int>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var ivp = line[j].Split(',');
                            bw.Write((ushort)ivp.Length);//长度
                            for (ushort i1 = 0; i1 < ivp.Length; i1++)
                            {
                                var pair = ivp[i1].Split('|');
                                if (pair.Length != 2) throw new Exception("Dictionary<int,int> parse error.Required format [int|int]");

                                var res = int.TryParse(pair[0].Trim(), out int key);
                                var res1 = int.TryParse(pair[1].Trim(), out int val);

                                if (res) bw.Write(key);
                                if (res1) bw.Write(val);
                            }
                        }
                    }
                }
                bw.Close();
            }
        }

        // A-Z AA-AZ BA-BZ etc.
        internal static string ConvertIntToOrderedLetter(int value)
        {
            var div = value / 26;
            var mod = value % 26;

            if (div > 0)
                return $"{((char)(div + 64))}{((char)(mod + 65))}";
            else
                return $"{((char)(mod + 65))}";
        }

        internal static bool IgnoreHeader(ExcelData.Header header)
        {
            // ignore del & enum~key
            return (header.define == TableHelper.HeaderFilter.k_DEL || header.define == TableHelper.HeaderFilter.k_ENUM_KEY);
        }

        internal static Dictionary<string, Type> s_TypeLut = new Dictionary<string, Type>
        {
            {"byte", typeof(byte) },
            {"int", typeof(int) },
            {"long", typeof(long) },
            {"float", typeof(float) },
            {"byte+", typeof(List<byte>) },
            {"int+", typeof(List<int>) },
            {"long+", typeof(List<long>) },
            {"float+", typeof(List<float>) },
            {"string", typeof(string) },
            {"[int|int]",typeof(Dictionary<int,int>) },

            // add more
        };

        internal static string[] InValidSheetNames = new string[]
        {
            "Sheet",
            "sheet"
        };

        internal class HeaderFilter
        {
            public const string k_DEL = "del";
            public const string k_CLIENT = "client";
            public const string k_SERVER = "server";
            public const string k_ENUM_KEY = "enum~key";
        }
    }
}
