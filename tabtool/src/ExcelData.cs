using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tabtool
{
    internal class ExcelData
    {
        public const int k_DataVersion = 1;

        public string tablName;
        public List<Header> header;
        public List<List<string>> rowValues;

        internal class Header
        {
            public string define;
            public string fieldComment;
            public string fieldName;
            public string fieldTypeName;
            //public ETableFieldType fieldType;

            public string GetCsharpTypeName()
            {
                //if (fieldType == ETableFieldType.Struct)
                //{
                //    return fieldTypeName;
                //}
                //if (fieldType == ETableFieldType.StructList)
                //{
                //    return string.Format("List<{0}>", fieldTypeName.Substring(0, fieldTypeName.Length - 1));
                //}
                return fieldTypeName;
            }
        }

        internal string GetClassName()
        {
            return "Cfg" + tablName;
        }

        internal string GetItemName()
        {
            return "Db" + tablName;
        }

        internal string GetEnumName()
        {
            return "EID" + tablName;
        }

        internal string ToString(bool ignore = false)
        {
            var sb = new StringBuilder(1024);
            //sb.AppendLine("Defines: ");
            sb.AppendLine($"Row Count: {rowValues.Count} Col Count: {rowValues[0].Count} {header.Count}");
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
