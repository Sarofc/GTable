using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Saro.Table
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
                    //Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{4}{ConvertIntToOrderedLetter(index)}]");
                    break;
                }

                cell = defineRow.GetCell(index);
                if (cell != null)
                {
                    data.header[index - 1].define = (cell.ToString());
                }
                else
                {
                    //Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{1}{ConvertIntToOrderedLetter(index)}]");
                }

                cell = commentRow.GetCell(index);
                if (cell != null)
                {
                    data.header[index - 1].fieldComment = (cell.ToString());
                }
                else
                {
                    //Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{2}{ConvertIntToOrderedLetter(index)}]");
                }

                cell = typeRow.GetCell(index);
                if (cell != null)
                {
                    data.header[index - 1].fieldTypeName = (cell.ToString());
                }
                else
                {
                    //Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{3}{ConvertIntToOrderedLetter(index)}]");
                }
            }

            cellNum = index - 1;

            while (data.header.Count > cellNum)
            {
                data.header.RemoveAt(data.header.Count - 1);
            }

            if (data.FieldNameDuplicated(out var duplicated))
            {
                throw new Exception("duplicate fieldName: " + duplicated);
            }

            //Console.WriteLine("CellNum: " + cellNum + "  NameRowNum: " + nameRow.LastCellNum);

            data.rowValues = new List<List<string>>(cellNum);
            for (int i = 4; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i) as XSSFRow;
                var keyCell = row.GetCell(1);
                if ((keyCell == null || string.IsNullOrEmpty(keyCell.ToString())))
                {
                    //Console.WriteLine($"key [{i + 1}B] is null or empty. break.");
                    break;
                }

                data.rowValues.Add(new List<string>(cellNum));

                for (int j = 1; j <= cellNum; j++)
                {
                    var cell = row.GetCell(j);

                    if (cell == null)
                    {
                        data.rowValues[i - 4].Add(string.Empty);
                        //Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{i + 1}{ConvertIntToOrderedLetter(i)}]");

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
                foreach (var name in s_InValidSheetNames)
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
        internal void WriteByteAsset(ExcelData data, string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                BinaryWriter bw = new BinaryWriter(fs, Encoding.UTF8);

                bw.Write(TableCfg.k_DataVersion);
                bw.Write(data.rowValues.Count);//行数据长度

                var keys = new List<int>();

                for (int i = 0; i < data.rowValues.Count; i++)
                {
                    var line = data.rowValues[i];

                    keys.Clear();

                    for (int j = 0; j < data.header.Count; j++)
                    {
                        if (IgnoreHeader(data.header[j])) continue;

                        if (!s_TypeLut.TryGetValue(data.header[j].fieldTypeName, out Type t))
                        {
                            throw new Exception("type is not support: " + data.header[j].fieldTypeName);
                        }

                        if (t == typeof(byte))
                        {
                            byte.TryParse(line[j], out byte val);
                            bw.Write(val);
                        }
                        else if (t == typeof(int))
                        {
                            var res = int.TryParse(line[j], out int val);
                            bw.Write(val);

                            if (IsKey(data.header[j]))
                            {
                                if (!res)
                                    throw new Exception($"write {t} failed.");

                                keys.Add(val);
                            }
                        }
                        else if (t == typeof(long))
                        {
                            var res = long.TryParse(line[j], out long val);
                            bw.Write(val);

                            if (!string.IsNullOrEmpty(line[j]) && !res)
                                throw new Exception($"write {t} failed.");
                        }
                        else if (t == typeof(float))
                        {
                            var res = float.TryParse(line[j], out float val);
                            bw.Write(val);

                            if (!string.IsNullOrEmpty(line[j]) && !res)
                                throw new Exception($"write {t} failed.");
                        }
                        else if (t == typeof(bool))
                        {
                            var res = bool.TryParse(line[j], out bool val);
                            bw.Write(val);

                            if (!string.IsNullOrEmpty(line[j]) && !res)
                                throw new Exception($"write {t} failed.");
                        }
                        else if (t == typeof(string))
                        {
                            bw.Write(line[j]);
                        }
                        else if (t == typeof(byte[]))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var arr = line[j].Split(',');
                            bw.Write((ushort)arr.Length);//长度
                            for (ushort i1 = 0; i1 < arr.Length; i1++)
                            {
                                var res = byte.TryParse(arr[i1].Trim(), out byte val);
                                bw.Write(val);

                                if (!string.IsNullOrEmpty(arr[j]) && !res)
                                    throw new Exception($"write {t} {arr[j]} failed.");
                            }
                        }
                        else if (t == typeof(int[]))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var arr = line[j].Split(',');
                            bw.Write((ushort)arr.Length);//长度
                            for (ushort i1 = 0; i1 < arr.Length; i1++)
                            {
                                var res = int.TryParse(arr[i1].Trim(), out int val);
                                bw.Write(val);

                                if (!string.IsNullOrEmpty(arr[j]) && !res)
                                    throw new Exception($"write {t} {arr[j]} failed.");
                            }
                        }
                        else if (t == typeof(long[]))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var arr = line[j].Split(',');
                            bw.Write((ushort)arr.Length);//长度
                            for (ushort i1 = 0; i1 < arr.Length; i1++)
                            {
                                var res = long.TryParse(arr[i1].Trim(), out long val);
                                bw.Write(val);

                                if (!string.IsNullOrEmpty(arr[j]) && !res)
                                    throw new Exception($"write {t} {arr[j]} failed.");
                            }
                        }
                        else if (t == typeof(float[]))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var arr = line[j].Split(',');
                            bw.Write((ushort)arr.Length);//长度
                            for (ushort i1 = 0; i1 < arr.Length; i1++)
                            {
                                var res = float.TryParse(arr[i1].Trim(), out float val);
                                bw.Write(val);

                                if (!string.IsNullOrEmpty(arr[j]) && !res)
                                    throw new Exception($"write {t} {arr[j]} failed.");
                            }
                        }
                        else if (t == typeof(Dictionary<int, int>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((ushort)0);
                                continue;
                            }
                            var arr = line[j].Split(',');
                            bw.Write((ushort)arr.Length);//长度
                            for (ushort i1 = 0; i1 < arr.Length; i1++)
                            {
                                var pair = arr[i1].Split('|');
                                if (pair.Length != 2) throw new Exception("Dictionary<int,int> parse error.Required format [int|int]");

                                var res = int.TryParse(pair[0].Trim(), out int key);
                                var res1 = int.TryParse(pair[1].Trim(), out int val);
                                bw.Write(key);
                                bw.Write(val);

                                if (!string.IsNullOrEmpty(pair[0]) && !res)
                                    throw new Exception($"write {t} {pair[0]} failed.");

                                if (!string.IsNullOrEmpty(pair[1]) && !res1)
                                    throw new Exception($"write {t} {pair[1]} failed.");
                            }
                        }
                    }

                    if (keys.Count < 1)
                    {
                        throw new Exception("at least 1 key");
                    }
                    else if (keys.Count > 4)
                    {
                        throw new Exception("more than 4 keys is not supportted");
                    }

                    ulong combinedKey = 0ul;
                    if (keys.Count == 1)
                        combinedKey = KeyHelper.GetKey(keys[0]);
                    else if (keys.Count == 2)
                        combinedKey = KeyHelper.GetKey(keys[0], keys[1]);
                    else if (keys.Count == 3)
                        combinedKey = KeyHelper.GetKey(keys[0], keys[1], keys[2]);
                    else if (keys.Count == 4)
                        combinedKey = KeyHelper.GetKey(keys[0], keys[1], keys[2], keys[3]);

                    bw.Write(combinedKey);
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

        internal static bool IsKey(ExcelData.Header header)
        {
            return string.CompareOrdinal(header.define, TableHelper.HeaderFilter.k_KEY) == 0;
        }

        internal static bool IgnoreHeader(ExcelData.Header header)
        {
            // ignore del & enum~key
            return string.CompareOrdinal(header.define, TableHelper.HeaderFilter.k_DEL) == 0 ||
                string.CompareOrdinal(header.define, TableHelper.HeaderFilter.k_ENUM_KEY) == 0;
        }

        internal static Dictionary<string, Type> s_TypeLut = new Dictionary<string, Type>
        {
            {"byte", typeof(byte) },
            {"int", typeof(int) },
            {"long", typeof(long) },
            {"float", typeof(float) },
            {"bool", typeof(bool) },
            {"byte+", typeof(byte[]) },
            {"int+", typeof(int[]) },
            {"long+", typeof(long[]) },
            {"float+", typeof(float[]) },
            {"string", typeof(string) },
            {"[int|int]",typeof(Dictionary<int,int>) },

            // add more
        };

        internal static string[] s_InValidSheetNames = new string[]
        {
            "Sheet",
            "sheet"
        };

        internal class HeaderFilter
        {
            public const string k_KEY = "key";
            public const string k_DEL = "del";
            public const string k_CLIENT = "client";
            public const string k_SERVER = "server";
            public const string k_ENUM_KEY = "enum~key";
            public const string k_Translate = "translate";
        }
    }
}
