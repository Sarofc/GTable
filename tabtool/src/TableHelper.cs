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
        internal void ParseFieldType(ExcelData.Header data, object obj)
        {
            var typeName = obj.ToString();

            if (typeName == "byte")
            {
                data.fieldType = (ETableFieldType.Byte);
            }
            else if (typeName == "int")
            {
                data.fieldType = (ETableFieldType.Int);
            }
            else if (typeName == "long")
            {
                data.fieldType = (ETableFieldType.Long);
            }
            else if (typeName == "float")
            {
                data.fieldType = (ETableFieldType.Float);
            }
            else if (typeName == "string")
            {
                data.fieldType = (ETableFieldType.String);
            }
            else if (typeName == "byte+")
            {
                data.fieldType = (ETableFieldType.ByteList);
            }
            else if (typeName == "int+")
            {
                data.fieldType = (ETableFieldType.IntList);
            }
            else if (typeName == "long+")
            {
                data.fieldType = (ETableFieldType.LongList);
            }
            else if (typeName == "float+")
            {
                data.fieldType = (ETableFieldType.FloatList);
            }
            //else if (typeName[typeName.Length - 1] == '+')
            //{
            //    data.fieldTypes.Add(ETableFieldType.StructList);
            //}
            //else
            //{
            //    data.fieldTypes.Add(ETableFieldType.Struct);
            //}
        }

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
            var colCount = nameRow.LastCellNum - 2;

            data.header = new List<ExcelData.Header>(colCount);
            for (int i = 0; i < colCount; i++)
            {
                data.header.Add(new ExcelData.Header());
            }

            // 以name为基准
            int cellNum = nameRow.LastCellNum;
            for (int i = 1; i < nameRow.LastCellNum; i++)
            {
                var cell = nameRow.GetCell(i);
                if (cell != null && !string.IsNullOrEmpty(cell.ToString()))
                {
                    data.header[i - 1].fieldName = (cell.ToString());
                }
                else
                {
                    cellNum = i - 1;
                    Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{3}{ConvertIntToOrderedLetter(i)}]");
                    break;
                }

                cell = defineRow.GetCell(i);
                if (cell != null)
                    data.header[i - 1].define = (cell.ToString());
                else
                    Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{0}{ConvertIntToOrderedLetter(i)}]");

                cell = commentRow.GetCell(i);
                if (cell != null)
                    data.header[i - 1].fieldComment = (cell.ToString());
                else
                    Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{1}{ConvertIntToOrderedLetter(i)}]");

                cell = typeRow.GetCell(i);
                if (cell != null)
                {
                    ParseFieldType(data.header[i - 1], cell.ToString());
                    data.header[i - 1].fieldTypeName = (cell.ToString());
                }
                else
                    Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{2}{ConvertIntToOrderedLetter(i)}]");
            }

            while (data.header.Count > cellNum)
            {
                data.header.RemoveAt(data.header.Count - 1);
            }

            data.rowValues = new List<List<string>>(cellNum);
            for (int i = 4; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i) as XSSFRow;
                var keyCell = row.GetCell(1);
                if ((keyCell == null || string.IsNullOrEmpty(keyCell.ToString())))
                {
                    Console.WriteLine($"key [{i}B] is null or empty. break.");
                    break;
                }

                data.rowValues.Add(new List<string>(cellNum));

                for (int j = 1; j <= cellNum; j++)
                {
                    var cell = row.GetCell(j);

                    if (cell == null)
                    {
                        data.rowValues[i - 4].Add(string.Empty);
                        Console.WriteLine($"Sheet [{sheet.SheetName}] cell is null at: [{0}{ConvertIntToOrderedLetter(i)}]");

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
                foreach (var name in TabToolConfig.InValidSheetNames)
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

        /// <summary>
        /// 获取 TableMeta
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        //internal TableMeta GetTableMeta(string filename, ExcelData data)
        //{
        //    TableMeta meta = new TableMeta();
        //    meta.TableName = Path.GetFileNameWithoutExtension(filename);
        //    for (int i = 1; i < data.fieldNames.Count; i++)
        //    {
        //        // TODO filter define
        //        if (data.defines[i] == TabToolConfig.ExportFilter.k_NOT) continue;

        //        TableField field = new TableField();
        //        field.commits = data.fieldCommits[i];
        //        field.fieldName = data.fieldNames[i];
        //        field.typeName = data.fieldTypeNames[i];
        //        if (field.typeName == "int") { field.fieldType = ETableFieldType.Int; }
        //        else if (field.typeName == "float") { field.fieldType = ETableFieldType.Float; }
        //        else if (field.typeName == "string") { field.fieldType = ETableFieldType.String; }
        //        else if (field.typeName == "int+") { field.fieldType = ETableFieldType.IntList; }
        //        else if (field.typeName == "float+") { field.fieldType = ETableFieldType.FloatList; }
        //        else if (field.typeName[field.typeName.Length - 1] == '+') { field.fieldType = ETableFieldType.StructList; }
        //        else { field.fieldType = ETableFieldType.Struct; }
        //        meta.Fields.Add(field);
        //    }

        //    return meta;
        //}

        /// <summary>
        /// 写入txt
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        internal void WriteTxtAsset(ExcelData data, string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);

                //for (int i = 0; i < data.value.Count; i++)
                //{
                //    var line = data.value[i];
                //    for (int j = 0; j < line.Count; j++)
                //    {
                //        // TODO filter define
                //        if (data.defines[j] == TabToolConfig.ExportFilter.k_NOT) continue;
                //        if (j == line.Count - 1)
                //        {
                //            sw.Write(line[j]);
                //        }
                //        else
                //        {
                //            sw.Write(line[j] + "\t");
                //        }
                //    }
                //    sw.WriteLine();
                //}
                sw.Close();
            }
        }

        /// <summary>
        /// 写入bytes
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="path"></param>
        internal void WriteByteAsset(ExcelData data, string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                BinaryWriter sw = new BinaryWriter(fs, Encoding.UTF8);

                sw.Write(ExcelData.k_DataVersion);
                //var typeList = new List<byte>();
                for (int i = 0; i < data.header.Count; i++)
                {
                    if (data.header[i].define == TabToolConfig.ExportFilter.k_NOT)
                        continue;
                    //typeList.Add((byte)data.header[i].fieldType);
                }

                //sw.Write(typeList.Count);
                //for (int i = 0; i < typeList.Count; i++)
                //{
                //    sw.Write(typeList[i]);
                //}

                sw.Write(data.rowValues.Count);//数据长度

                for (int i = 0; i < data.rowValues.Count; i++)
                {
                    var line = data.rowValues[i];
                    for (int j = 0; j < line.Count; j++)
                    {
                        // TODO filter define
                        if (data.header[j].define == TabToolConfig.ExportFilter.k_NOT) continue;

                        switch (data.header[j].fieldType)
                        {
                            case ETableFieldType.Byte:
                                byte.TryParse(line[j], out byte bv);
                                sw.Write(bv);
                                break;
                            case ETableFieldType.Int:
                                int.TryParse(line[j], out int iv);
                                sw.Write(iv);
                                break;
                            case ETableFieldType.Long:
                                long.TryParse(line[j], out long lv);
                                sw.Write(lv);
                                break;
                            case ETableFieldType.Float:
                                float.TryParse(line[j], out float fv);
                                sw.Write(fv);
                                break;
                            case ETableFieldType.String:
                                sw.Write(line[j]);
                                break;
                            case ETableFieldType.ByteList:
                                var bvs = line[j].Split(',');
                                sw.Write(bvs.Length);//长度
                                for (int i1 = 0; i1 < bvs.Length; i1++)
                                {
                                    byte.TryParse(bvs[i1].Trim(), out byte bvl);
                                    sw.Write(bvl);
                                }
                                break;
                            case ETableFieldType.IntList:
                                var ivs = line[j].Split(',');
                                sw.Write(ivs.Length);//长度
                                for (int i1 = 0; i1 < ivs.Length; i1++)
                                {
                                    int.TryParse(ivs[i1].Trim(), out int ivl);
                                    sw.Write(ivl);
                                }
                                break;
                            case ETableFieldType.LongList:
                                var lvs = line[j].Split(',');
                                sw.Write(lvs.Length);//长度
                                for (int i1 = 0; i1 < lvs.Length; i1++)
                                {
                                    long.TryParse(lvs[i1].Trim(), out long lvl);
                                    sw.Write(lvl);
                                }
                                break;
                            case ETableFieldType.FloatList:
                                var fvs = line[j].Split(',');
                                sw.Write(fvs.Length);//长度
                                for (int i1 = 0; i1 < fvs.Length; i1++)
                                {
                                    float.TryParse(fvs[i1].Trim(), out float fvl);
                                    sw.Write(fvl);
                                }
                                break;
                            //case ETableFieldType.Struct:
                            //    break;
                            //case ETableFieldType.StructList:
                            //    break;
                            default:
                                break;
                        }
                    }
                }
                sw.Close();
            }
        }

        // A-Z AA-AZ BA-BZ etc.
        public string ConvertIntToOrderedLetter(int value)
        {
            var div = value / 26;
            var mod = value % 26;

            if (div > 0)
                return $"{((char)(div + 64))}{((char)(mod + 65))}";
            else
                return $"{((char)(mod + 65))}";
        }

        internal static Dictionary<ETableFieldType, Type> s_TypeLut = new Dictionary<ETableFieldType, Type>
        {
            {ETableFieldType.Byte, typeof(byte) },
            {ETableFieldType.Int, typeof(int) },
            {ETableFieldType.Long, typeof(long) },
            {ETableFieldType.Float, typeof(float) },
            {ETableFieldType.ByteList, typeof(List<byte>) },
            {ETableFieldType.IntList, typeof(List<int>) },
            {ETableFieldType.LongList, typeof(List<long>) },
            {ETableFieldType.FloatList  , typeof(List<float>) },
            {ETableFieldType.String, typeof(string) },
        };
    }
}
