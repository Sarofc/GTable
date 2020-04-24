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
using System.Runtime.Serialization.Formatters.Binary;

namespace tabtool
{
    public class TableHelper
    {
        /// <summary>
        /// 根据sheet，实例化DataTable
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        internal DataTable GetDataTable(ISheet sheet, bool toBytes = false)
        {
            IEnumerator rows = sheet.GetRowEnumerator();

            DataTable dt = new DataTable();
            IRow typeRow = sheet.GetRow(2);
            IRow nameRow = sheet.GetRow(3);
            for (int j = nameRow.FirstCellNum; j < (nameRow.LastCellNum); j++)
            {
                dt.Columns.Add(j.ToString());
            }

            var dataReader = new DataReader();

            while (rows.MoveNext())
            {
                IRow row = (XSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);
                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        if (toBytes)
                        {
                            var typeName = typeRow.GetCell(i).ToString();
                            if (typeName == "int")
                            {
                                dr[i] = dataReader.GetInt(cell.ToString());
                            }
                            else if (typeName == "float")
                            {
                                dr[i] = dataReader.GetFloat(cell.ToString());
                            }
                            else if (typeName == "string")
                            {
                                dr[i] = cell.ToString();
                            }
                            else if (typeName == "int+")
                            {
                                dr[i] = dataReader.GetIntList(cell.ToString());
                            }
                            else if (typeName == "float+")
                            {
                                dr[i] = dataReader.GetFloatList(cell.ToString());
                            }
                            //else if (typeRow.GetCell(i).ToString() == "string+")
                            //{
                            //    dr[i] = dataReader.GetStringList(cell.ToString());
                            //}
                            else if (typeName[typeName.Length - 1] == '+')
                            {
                                dr[i] = cell.ToString();//dataReader.GetObject(cell.ToString(), Type.GetType(typeName));
                            }
                            else
                            {
                                dr[i] = cell.ToString(); //dataReader.GetObjectList(cell.ToString(), Type.GetType(typeName));
                            }
                        }
                        else
                        {
                            dr[i] = cell.ToString();
                        }
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        /// <summary>
        /// 加载Excel表里的所有可用Sheet
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        internal List<ISheet> ImportExcelFile(string filePath)
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

                return true;
            }

            return false;
        }

        /// <summary>
        /// 获取 TableMeta
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        internal TableMeta GetTableMeta(string filename, DataTable dt)
        {
            TableMeta meta = new TableMeta();
            meta.TableName = Path.GetFileNameWithoutExtension(filename);
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                // TODO filter define
                if (dt.Rows[0].ItemArray[i].ToString() == TabToolConfig.ExportFilter.k_NOT) continue;

                TableField field = new TableField();
                field.commits = dt.Rows[1].ItemArray[i].ToString();
                field.fieldName = dt.Rows[2].ItemArray[i].ToString();
                field.typeName = dt.Rows[3].ItemArray[i].ToString();
                if (field.typeName == "int") { field.fieldType = ETableFieldType.Int; }
                else if (field.typeName == "float") { field.fieldType = ETableFieldType.Float; }
                else if (field.typeName == "string") { field.fieldType = ETableFieldType.String; }
                else if (field.typeName == "int+") { field.fieldType = ETableFieldType.IntList; }
                else if (field.typeName == "float+") { field.fieldType = ETableFieldType.FloatList; }
                else if (field.typeName == "string+") { field.fieldType = ETableFieldType.StringList; }
                else if (field.typeName[field.typeName.Length - 1] == '+') { field.fieldType = ETableFieldType.StructList; }
                else { field.fieldType = ETableFieldType.Struct; }
                meta.Fields.Add(field);
            }

            return meta;
        }

        /// <summary>
        /// 写入txt
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="path"></param>
        internal void WriteTxtAsset(DataTable dt, string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);

                for (int i1 = 0; i1 < dt.Rows.Count; i1++)
                {
                    // key == not 直接跳过整个表 应该写在此方法外面
                    //if (dt.Rows[0].ItemArray[1].ToString() == TabToolConfig.ExportFilter.k_NOT) break;

                    if (i1 == 0 || i1 == 1 || /*i == 2 ||*/ i1 == 3) continue;//只保留名称
                    for (int j = 1; j < dt.Columns.Count; j++)
                    {
                        // TODO filter define
                        if (dt.Rows[0].ItemArray[j].ToString() == TabToolConfig.ExportFilter.k_NOT) continue;
                        if (j == dt.Columns.Count - 1)
                        {
                            sw.Write(dt.Rows[i1].ItemArray[j].ToString());
                        }
                        else
                        {
                            sw.Write(dt.Rows[i1].ItemArray[j].ToString() + "\t");
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
        /// <param name="dt"></param>
        /// <param name="path"></param>
        internal void WriteByteAsset(DataTable dt, string path)
        {
            byte[] binaryDataResult = null;
            using (MemoryStream memStream = new MemoryStream())
            {
                BinaryFormatter brFormatter = new BinaryFormatter();
                dt.RemotingFormat = SerializationFormat.Binary;
                brFormatter.Serialize(memStream, dt);
                binaryDataResult = memStream.ToArray();
            }

            File.WriteAllBytes(path, binaryDataResult);
        }

    }

    //表字段读取
    class DataReader
    {
        public List<string> GetStringList(string s, char delim)
        {
            string[] t = s.Split(delim);
            List<string> ret = new List<string>();
            ret.AddRange(t);
            return ret;
        }

        public int GetInt(string s)
        {
            return int.Parse(s);
        }

        public List<int> GetIntList(string s)
        {
            string[] vs = s.Split(',');
            List<int> ret = new List<int>();
            foreach (var ss in vs)
            {
                int x = int.Parse(ss);
                ret.Add(x);
            }
            return ret;
        }

        public float GetFloat(string s)
        {
            return float.Parse(s);
        }

        public List<float> GetFloatList(string s)
        {
            string[] vs = s.Split(',');
            List<float> ret = new List<float>();
            foreach (var ss in vs)
            {
                float x = float.Parse(ss);
                ret.Add(x);
            }
            return ret;
        }

        //public T GetObject<T>(string s) where T : ITableObject, new()
        //{
        //    T obj = new T();
        //    obj.FromString(s);
        //    return obj;
        //}

        //public List<T> GetObjectList<T>(string s) where T : ITableObject, new()
        //{
        //    string[] vs = s.Split(';');
        //    List<T> ret = new List<T>();
        //    foreach (var ss in vs)
        //    {
        //        ret.Add(GetObject<T>(ss));
        //    }
        //    return ret;
        //}
    };
}
