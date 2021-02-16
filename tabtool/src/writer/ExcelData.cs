using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Saro.Table
{
    internal class ExcelData
    {
        public const int k_DataVersion = 1;

        public string tablName;
        public List<Header> header;
        public List<List<string>> rowValues;

        private static HashSet<string> s_Set = new HashSet<string>();

        internal int GetKeyCount() => header.Count(h => TableHelper.IsKey(h));

        /// <summary>
        /// 字段名是否重复
        /// </summary>
        /// <returns></returns>
        internal bool FieldNameDuplicated()
        {
            s_Set.Clear();

            foreach (var h in header)
            {
                if (s_Set.Contains(h.fieldName)) return true;
                s_Set.Add(h.fieldName);
            }

            return false;
        }

        /// <summary>
        /// 数据表头
        /// </summary>
        internal class Header
        {
            /// <summary>
            /// 定义
            /// </summary>
            public string define;
            /// <summary>
            /// 注释
            /// </summary>
            public string fieldComment;
            /// <summary>
            /// 字段名，唯一
            /// </summary>
            public string fieldName;
            /// <summary>
            /// 字段类型
            /// </summary>
            public string fieldTypeName;
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
