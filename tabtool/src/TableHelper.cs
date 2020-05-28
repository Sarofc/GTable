using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using NPOI.XSSF.UserModel;
using System.Xml;
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
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);

                for (int i = 0; i < data.rowValues.Count; i++)
                {
                    var line = data.rowValues[i];
                    for (int j = 0; j < line.Count; j++)
                    {
                        if (IgnoreHeader(data.header[i])) continue;
                        if (j == line.Count - 1)
                        {
                            sw.Write(line[j]);
                        }
                        else
                        {
                            sw.Write(line[j] + "\t");
                        }
                    }
                    sw.WriteLine();
                }
                sw.Close();
            }
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
                BinaryWriter sw = new BinaryWriter(fs, Encoding.UTF8);

                sw.Write(ExcelData.k_DataVersion);
                sw.Write(data.rowValues.Count);//行数据长度

                for (int i = 0; i < data.rowValues.Count; i++)
                {
                    var line = data.rowValues[i];
                    for (int j = 0; j < line.Count; j++)
                    {
                        if (IgnoreHeader(data.header[j])) continue;

                        if (!s_TypeLut.TryGetValue(data.header[j].fieldTypeName, out Type t))
                        {
                            throw new Exception("type is not support: " + data.header[j].fieldTypeName);
                        }

                        if (t == typeof(byte))
                        {
                            byte.TryParse(line[j], out byte bv);
                            sw.Write(bv);
                        }
                        else if (t == typeof(int))
                        {
                            int.TryParse(line[j], out int iv);
                            sw.Write(iv);
                        }
                        else if (t == typeof(long))
                        {
                            long.TryParse(line[j], out long lv);
                            sw.Write(lv);
                        }
                        else if (t == typeof(float))
                        {
                            float.TryParse(line[j], out float fv);
                            sw.Write(fv);
                        }
                        else if (t == typeof(string))
                        {
                            sw.Write(line[j]);
                        }
                        else if (t == typeof(List<byte>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                sw.Write((ushort)0);
                                continue;
                            }
                            var bvs = line[j].Split(',');
                            sw.Write((ushort)bvs.Length);//长度
                            for (ushort i1 = 0; i1 < bvs.Length; i1++)
                            {
                                byte.TryParse(bvs[i1].Trim(), out byte bvl);
                                sw.Write(bvl);
                            }
                        }
                        else if (t == typeof(List<int>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                sw.Write((ushort)0);
                                continue;
                            }
                            var ivs = line[j].Split(',');
                            sw.Write((ushort)ivs.Length);//长度
                            for (ushort i1 = 0; i1 < ivs.Length; i1++)
                            {
                                int.TryParse(ivs[i1].Trim(), out int ivl);
                                sw.Write(ivl);
                            }
                        }
                        else if (t == typeof(List<long>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                sw.Write((ushort)0);
                                continue;
                            }
                            var lvs = line[j].Split(',');
                            sw.Write((ushort)lvs.Length);//长度
                            for (ushort i1 = 0; i1 < lvs.Length; i1++)
                            {
                                long.TryParse(lvs[i1].Trim(), out long lvl);
                                sw.Write(lvl);
                            }
                        }
                        else if (t == typeof(List<float>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                sw.Write((ushort)0);
                                continue;
                            }
                            var fvs = line[j].Split(',');
                            sw.Write((ushort)fvs.Length);//长度
                            for (ushort i1 = 0; i1 < fvs.Length; i1++)
                            {
                                float.TryParse(fvs[i1].Trim(), out float fvl);
                                sw.Write(fvl);
                            }
                            break;
                        }
                    }
                }
                sw.Close();
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
