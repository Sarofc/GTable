using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Saro.Table
{
    public class ExcelData
    {
        public string tablName;
        public List<Header> header;
        public List<List<string>> rowValues;

        private readonly HashSet<string> m_Set = new HashSet<string>(40);

        /// <summary>
        /// key的个数
        /// </summary>
        /// <returns></returns>
        internal int GetKeyCount() => header.Count(h => TableHelper.IsKey(h));

        /// <summary>
        /// key的字段名称
        /// </summary>
        /// <returns></returns>
        internal List<string> GetKeyNames()
        {
            var ret = new List<string>();
            foreach (var h in header)
            {
                if (TableHelper.IsKey(h))
                {
                    ret.Add(h.fieldName);
                }
            }
            return ret;
        }

        /// <summary>
        /// 字段名是否重复
        /// </summary>
        /// <returns></returns>
        internal bool FieldNameDuplicated(out string duplicatedString)
        {
            m_Set.Clear();

            duplicatedString = null;

            foreach (var h in header)
            {
                if (TableHelper.IgnoreHeader(h)) continue;

                if (m_Set.Contains(h.fieldName))
                {
                    duplicatedString = h.fieldName;
                    return true;
                }
                m_Set.Add(h.fieldName);
            }

            return false;
        }

        /// <summary>
        /// 数据表头
        /// </summary>
        public class Header
        {
            /// <summary>
            /// 定义
            /// </summary>
            public string define { get => metas[0]; set => metas[0] = value; }
            /// <summary>
            /// 注释
            /// </summary>
            public string fieldComment { get => metas[1]; set => metas[1] = value; }
            /// <summary>
            /// 字段类型
            /// </summary>
            public string fieldTypeName { get => metas[2]; set => metas[2] = value; }
            /// <summary>
            /// 字段名，唯一
            /// </summary>
            public string fieldName { get => metas[3]; set => metas[3] = value; }

            public string[] metas = new string[4];
        }

        /// <summary>
        /// wrapper类名
        /// </summary>
        /// <returns></returns>
        internal string GetWrapperClassName()
        {
            return "csv" + tablName;
        }

        /// <summary>
        /// 数据实体类名
        /// </summary>
        /// <returns></returns>
        internal string GetEntityClassName()
        {
            return "t" + tablName;
        }

        /// <summary>
        /// 枚举名
        /// </summary>
        /// <returns></returns>
        internal string GetEnumName()
        {
            return "e" + tablName;
        }

        /// <summary>
        /// debug用，打印数据表
        /// </summary>
        /// <param name="ignore"></param>
        /// <returns></returns>
        internal string ToString(bool ignore = false)
        {
            var sb = new StringBuilder(1024);
            //sb.AppendLine("Defines: ");
            sb.AppendLine($"Row Count: {rowValues.Count} Col Count: {rowValues[0].Count} Head Count: {header.Count}");
            foreach (var data in header)
            {
                if (ignore && TableHelper.IgnoreHeader(data)) continue;
                sb.Append(data.define).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Commits: ");
            foreach (var data in header)
            {
                if (ignore && TableHelper.IgnoreHeader(data)) continue;
                sb.Append(data.fieldComment).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Types: ");
            foreach (var data in header)
            {
                if (ignore && TableHelper.IgnoreHeader(data)) continue;
                sb.Append(data.fieldTypeName).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Names: ");
            foreach (var data in header)
            {
                if (ignore && TableHelper.IgnoreHeader(data)) continue;
                sb.Append(data.fieldName).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Data: ");
            for (int i1 = 0; i1 < rowValues.Count; i1++)
            {
                var line = rowValues[i1];
                var count = -1;
                foreach (var word in line)
                {
                    count++;
                    if (ignore && TableHelper.IgnoreHeader(header[count])) continue;
                    sb.Append(word).Append("\t");
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }
    }
}
