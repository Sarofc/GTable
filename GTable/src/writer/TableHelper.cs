
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using ExcelDataReader;
using System.Linq;
using System.Diagnostics;

namespace Saro.Table
{
    public static class TableHelper
    {
        static TableHelper()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private static ExcelData ParseExcel(IExcelDataReader reader)
        {
            // excel 真实行列数
            var colCount = reader.FieldCount; // 减少1 = 实际数据 列
            var rowCount = reader.RowCount; // 要减少4 = 实际数据 行

            if (!reader.Read() || colCount < 3 || rowCount < 5)
            {
                throw new Exception($"invalid sheet: {reader.Name}. colCount < 2 or rowCount < 5 or other error");
            }

            var data = new ExcelData();

            data.tablName = reader.Name;
            data.header = new List<ExcelData.Header>(colCount);
            data.rowValues = new List<List<string>>(rowCount - 4);

            //Console.WriteLine($"sheet: {reader.Name}  row x col = {rowCount} x {colCount}");

            for (int i = 1, n = colCount; i < n; i++)
            {
                data.header.Add(new ExcelData.Header());
            }

            int rowIndex = 0;
            do
            {
                if (rowIndex <= 3)
                {
                    for (int i = 1, n = colCount; i < n; i++)
                    {
                        var isNull = reader.IsDBNull(i);
                        if (rowIndex == 3)
                        {
                            if (isNull)
                            {
                                colCount = i;
                                //Console.WriteLine($"{data.tablName} typeName is empty. skip [{rowIndex + 1}{ConvertIntToOrderedLetter(i)}]. colCount: {colCount} i: {i}");
                                break;
                            }
                        }

                        var value = isNull ? string.Empty : reader.GetValue(i).ToString();
                        data.header[i - 1].metas[rowIndex] = value;

                        //Console.WriteLine($"[{rowIndex + 1}{ConvertIntToOrderedLetter(i)}] = {value}");
                    }

                    //Console.WriteLine($"{data.header.Count} {string.Join("\t", data.header.Select(h => h.metas[rowIndex]))}");
                }
                else
                {
                    for (int i = 1, n = colCount; i < n; i++)
                    {
                        var isNull = reader.IsDBNull(i);

                        if (i == 1)
                        {
                            if (isNull)
                            {
                                //Console.WriteLine($"{data.tablName} key is empty. skip: [{rowIndex + 1}{ConvertIntToOrderedLetter(i)}]");
                                // 第一个 key 没有值，则中断处理
                                goto END;
                            }
                            data.rowValues.Add(new List<string>(colCount - 1));
                        }

                        var value = isNull ? string.Empty : reader.GetValue(i).ToString();
                        data.rowValues[rowIndex - 4].Add(value);

                        //Console.WriteLine($"[{rowIndex + 1}{ConvertIntToOrderedLetter(i)}] = {value}");
                    }
                }

                ++rowIndex;
            }
            while (reader.Read());

        END:


            while (data.header.Count > colCount - 1)
            {
                data.header.RemoveAt(data.header.Count - 1);
            }

            //Console.WriteLine($"sheet: {reader.Name} row x col = {rowCount} x {colCount}");

            Debug.Assert(colCount - 1 == data.header.Count, $"{data.tablName}'s colCount is invalid.");
            Debug.Assert(rowCount - 4 == data.rowValues.Count, $"{data.tablName}'s rowCount is invalid.");

            return data;
        }

        public static IEnumerable<ExcelData> LoadExcel(string filePath)
        {
            //Console.WriteLine("process xls: " + filePath);

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var counter = 0;
                    do
                    {
                        if (IsSheetNameValid(reader.Name))
                        {
                            counter++;
                            ExcelData data;
                            try
                            {
                                Console.WriteLine($"parsing...... {reader.Name} ({counter}/{reader.ResultsCount})");
                                data = ParseExcel(reader);
                            }
                            catch (Exception e)
                            {
                                throw new Exception($"excel:{filePath} sheet:{reader.Name} 读取失败.", e);
                            }
                            if (data != null)
                            {
                                yield return data;
                            }
                        }
                    }
                    while (reader.NextResult());
                }
            }
        }

        internal static bool IsSheetNameValid(string sheetName)
        {
            if (!string.IsNullOrEmpty(sheetName))
            {
                if (sheetName.StartsWith("~"))
                {
                    return false;
                }

                foreach (var name in s_InValidSheetNames)
                {
                    if (sheetName.Contains(name))
                    {
                        return false;
                    }
                }

                return true;
            }

            return false;
        }

        /// <summary>
        /// 写入bytes
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        internal static void WriteByteAsset(ExcelData data, string path)
        {
            //Console.WriteLine($"{data.tablName} write to {path}");

            const int MAX_ARRAY_LENGTH = 64;

            Span<byte> bytes = stackalloc byte[MAX_ARRAY_LENGTH];
            Span<int> ints = stackalloc int[MAX_ARRAY_LENGTH];
            Span<long> longs = stackalloc long[MAX_ARRAY_LENGTH];
            Span<float> floats = stackalloc float[MAX_ARRAY_LENGTH];

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
                            throw new Exception($"type is not support. {data.tablName} col: [{j}] typeName: {data.header[j].fieldTypeName} define: {data.header[j].define}");
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
                                    throw new Exception($"{data.tablName} write {t} \"{line[j]}\" failed.");

                                keys.Add(val);
                            }
                        }
                        else if (t == typeof(long))
                        {
                            var res = long.TryParse(line[j], out long val);
                            bw.Write(val);

                            if (!string.IsNullOrEmpty(line[j]) && !res)
                                throw new Exception($"{data.tablName} write {t} \"{line[j]}\" failed.");
                        }
                        else if (t == typeof(float))
                        {
                            var res = float.TryParse(line[j], out float val);
                            bw.Write(val);

                            if (!string.IsNullOrEmpty(line[j]) && !res)
                                throw new Exception($"{data.tablName} write {t} \"{line[j]}\" failed.");
                        }
                        else if (t == typeof(bool))
                        {
                            var res = bool.TryParse(line[j], out bool val);
                            bw.Write(val);

                            if (!string.IsNullOrEmpty(line[j]) && !res)
                                throw new Exception($"{data.tablName} write {t} \"{line[j]}\" failed.");
                        }
                        else if (t == typeof(string))
                        {
                            bw.Write(line[j]);
                        }
                        else if (t == typeof(byte[]))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((byte)0);
                                continue;
                            }

                            var span = line[j].AsSpan();
                            var arr = span.Split(',');
                            bytes.Clear();
                            byte index = 0;
                            foreach (var item in arr)
                            {
                                var chrSpan = span[item].Trim();

                                if (!byte.TryParse(chrSpan, out var val))
                                {
                                    if (!chrSpan.IsEmpty)
                                        throw new Exception($"{data.tablName} write {t} {chrSpan.ToString()} failed.");
                                }
                                bytes[index++] = val;
                            }

                            bw.Write(index);
                            for (int k = 0; k < index; k++)
                            {
                                bw.Write(bytes[k]);
                            }
                        }
                        else if (t == typeof(int[]))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((byte)0);
                                continue;
                            }

                            var span = line[j].AsSpan();
                            var arr = span.Split(',');
                            ints.Clear();
                            byte index = 0;
                            foreach (var item in arr)
                            {
                                var chrSpan = span[item].Trim();

                                if (!int.TryParse(chrSpan, out var val))
                                {
                                    if (!chrSpan.IsEmpty)
                                        throw new Exception($"{data.tablName} write {t} {chrSpan.ToString()} failed.");
                                }
                                ints[index++] = val;
                            }

                            bw.Write(index);
                            for (int k = 0; k < index; k++)
                            {
                                bw.Write(ints[k]);
                            }
                        }
                        else if (t == typeof(long[]))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((byte)0);
                                continue;
                            }

                            var span = line[j].AsSpan();
                            var arr = span.Split(',');
                            longs.Clear();
                            byte index = 0;
                            foreach (var item in arr)
                            {
                                var chrSpan = span[item].Trim();

                                if (!long.TryParse(chrSpan, out var val))
                                {
                                    if (!chrSpan.IsEmpty)
                                        throw new Exception($"{data.tablName} write {t} {chrSpan.ToString()} failed.");
                                }
                                longs[index++] = val;
                            }

                            bw.Write(index);
                            for (int k = 0; k < index; k++)
                            {
                                bw.Write(longs[k]);
                            }
                        }
                        else if (t == typeof(float[]))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((byte)0);
                                continue;
                            }

                            var span = line[j].AsSpan();
                            var arr = span.Split(',');
                            byte index = 0;
                            floats.Clear();
                            foreach (var item in arr)
                            {
                                var chrSpan = span[item].Trim();

                                if (!float.TryParse(chrSpan, out var val))
                                {
                                    if (!chrSpan.IsEmpty)
                                        throw new Exception($"{data.tablName} write {t} {chrSpan.ToString()} failed.");
                                }
                                floats[index++] = val;
                            }

                            bw.Write(index);
                            for (int k = 0; k < index; k++)
                            {
                                bw.Write(floats[k]);
                            }
                        }
                        else if (t == typeof(Dictionary<int, int>))
                        {
                            if (string.IsNullOrEmpty(line[j]))
                            {
                                bw.Write((byte)0);
                                continue;
                            }

                            var span = line[j].AsSpan();
                            var arr = span.Split(',');
                            ints.Clear();
                            byte index = 0;
                            foreach (var item in arr)
                            {
                                var chrSpan = span[item];
                                var splitIndex = chrSpan.IndexOf('|');

                                if (splitIndex < 0) throw new Exception("Dictionary<int,int> parse error. Required format int|int. eg. 1|2");

                                var pair0 = chrSpan.Slice(0, splitIndex).Trim();
                                var pair1 = chrSpan.Slice(splitIndex + 1).Trim();

                                if (!int.TryParse(pair0, out var val0))
                                {
                                    if (!pair0.IsEmpty)
                                        throw new Exception($"write {t} {pair0.ToString()} failed.");
                                }
                                ints[index++] = val0;

                                if (!int.TryParse(pair1, out var val1))
                                {
                                    if (!pair1.IsEmpty)
                                        throw new Exception($"write {t} {pair1.ToString()} failed.");
                                }
                                ints[index++] = val1;
                            }

                            bw.Write((byte)(index / 2));
                            for (int k = 0; k < index; k++)
                            {
                                bw.Write(ints[k]);
                            }
                        }
                    }

                    if (keys.Count < 1)
                    {
                        throw new Exception($"{data.tablName} at least 1 key");
                    }
                    else if (keys.Count > 4)
                    {
                        throw new Exception($"{data.tablName} more than 4 keys is not supportted");
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
