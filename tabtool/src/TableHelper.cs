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

            if (typeName == "int")
            {
                data.fieldType = (ETableFieldType.Int);
            }
            else if (typeName == "float")
            {
                data.fieldType = (ETableFieldType.Float);
            }
            else if (typeName == "string")
            {
                data.fieldType = (ETableFieldType.String);
            }
            else if (typeName == "int+")
            {
                data.fieldType = (ETableFieldType.IntList);
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
            var colCount = nameRow.LastCellNum - 1;

            data.header = new List<ExcelData.Header>(colCount);
            for (int i = 0; i < colCount; i++)
            {
                data.header.Add(new ExcelData.Header());
            }

            // 以name为基准
            int cellNum = nameRow.LastCellNum;
            for (int i = 0; i < nameRow.LastCellNum; i++)
            {
                var cell = nameRow.GetCell(i);
                if (cell != null && !string.IsNullOrEmpty(cell.ToString()))
                {
                    data.header[i].fieldName = (cell.ToString());
                }
                else
                {
                    cellNum = i;
                    Console.WriteLine($"Sheet {sheet.SheetName} cell is null at: [{2}{ConvertIntToLetter(i)}]");
                    break;
                }

                cell = defineRow.GetCell(i);
                if (cell != null)
                    data.header[i].define = (cell.ToString());
                else
                    Console.WriteLine($"Sheet {sheet.SheetName} cell is null at: [{0}{ConvertIntToLetter(i)}]");

                cell = commentRow.GetCell(i);
                if (cell != null)
                    data.header[i].fieldComment = (cell.ToString());
                else
                    Console.WriteLine($"Sheet {sheet.SheetName} cell is null at: [{0}{ConvertIntToLetter(i)}]");

                cell = typeRow.GetCell(i);
                if (cell != null)
                {
                    ParseFieldType(data.header[i], cell.ToString());
                    data.header[i].fieldTypeName = (cell.ToString());
                }
                else
                    Console.WriteLine($"Sheet {sheet.SheetName} cell is null at: [{0}{ConvertIntToLetter(i)}]");
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
                    Console.WriteLine("key is null or empty. break.");
                    break;
                }

                data.rowValues.Add(new List<string>(cellNum));

                for (int j = 1; j < cellNum; j++)
                {
                    var cell = row.GetCell(j);

                    if (cell == null)
                    {
                        data.rowValues[i - 4].Add(string.Empty);
                        Console.WriteLine($"Sheet {sheet.SheetName} cell is null at: [{0}{ConvertIntToLetter(i)}]");

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

        //TODO
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

                //sw.Write(ExcelData.k_DataVersion);
                //var typeList = new List<byte>();
                //for (int i = 0; i < data.fieldTypes.Count; i++)
                //{
                //    if (data.defines[i] == TabToolConfig.ExportFilter.k_NOT)
                //        continue;
                //    typeList.Add((byte)data.fieldTypes[i]);
                //}

                //sw.Write(typeList.Count);
                //for (int i = 0; i < typeList.Count; i++)
                //{
                //    sw.Write(typeList[i]);
                //}

                //sw.Write(data.value.Count);//数据长度

                //for (int i = 0; i < data.value.Count; i++)
                //{
                //    var line = data.value[i];
                //    for (int j = 0; j < line.Count; j++)
                //    {
                //        // TODO filter define
                //        if (data.defines[j] == TabToolConfig.ExportFilter.k_NOT) continue;

                //        switch (data.fieldTypes[j])
                //        {
                //            case ETableFieldType.Int:
                //                int.TryParse(line[j], out int iv);
                //                sw.Write(iv);
                //                break;
                //            case ETableFieldType.Float:
                //                float.TryParse(line[j], out float fv);
                //                sw.Write(fv);
                //                break;
                //            case ETableFieldType.String:
                //                sw.Write(line[j]);
                //                break;
                //            case ETableFieldType.IntList:
                //                var ivs = line[j].Split(',');
                //                sw.Write(ivs.Length);//长度
                //                for (int i1 = 0; i1 < ivs.Length; i1++)
                //                {
                //                    int.TryParse(ivs[i1], out int ivl);
                //                    sw.Write(ivl);
                //                }
                //                break;
                //            case ETableFieldType.FloatList:
                //                var fvs = line[j].Split(',');
                //                sw.Write(fvs.Length);//长度
                //                for (int i1 = 0; i1 < fvs.Length; i1++)
                //                {
                //                    float.TryParse(fvs[i1], out float fvl);
                //                    sw.Write(fvl);
                //                }
                //                break;
                //            //case ETableFieldType.Struct:
                //            //    break;
                //            //case ETableFieldType.StructList:
                //            //    break;
                //            default:
                //                break;
                //        }
                //    }
                //}
                sw.Close();
            }
        }

        public string ConvertIntToLetter(int value)
        {
            var count = value / 26;
            var res = value % 26;
            
            if (count > 0)
                return $"{((char)(count + 64))}{((char)(res + 65))}";
            else
                return $"{((char)(res + 65))}";
        }
    }
}
