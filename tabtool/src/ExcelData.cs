using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tabtool
{
    internal enum ETableFieldType : byte
    {
        Byte,
        Int,
        Long,
        Float,
        String,
        Struct,// TODO
        ByteList,
        IntList,
        LongList,
        FloatList,
        StructList,// TODO
    }

    internal class ExcelData
    {
        public const int k_DataVersion = 1;

        public string tablName;
        public List<Header> header;
        public List<List<string>> rowValues;

        public class Header
        {
            public string define;
            public string fieldComment;
            public string fieldName;
            public string fieldTypeName;
            public ETableFieldType fieldType;

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

        public string GetClassName()
        {
            return "Cfg" + tablName + "Table";
        }

        public string GetItemName()
        {
            return "Tbs" + tablName + "Table";
        }

        public override string ToString()
        {
            var sb = new StringBuilder(1024);
            //sb.AppendLine("Defines: ");
            sb.AppendLine($"Row Count: {rowValues.Count} Col Count: {header.Count}");
            foreach (var data in header)
            {
                sb.Append(data.define).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Commits: ");
            foreach (var data in header)
            {
                sb.Append(data.fieldComment).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Types: ");
            foreach (var data in header)
            {
                sb.Append(data.fieldType).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Names: ");
            foreach (var data in header)
            {
                sb.Append(data.fieldName).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Data: ");
            for (int i1 = 0; i1 < rowValues.Count; i1++)
            {
                var line = rowValues[i1];
                foreach (var word in line)
                {
                    sb.Append(word).Append("\t");
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }
    }
}
