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

        public List<string> defines;
        public List<string> fieldCommits;
        public List<string> fieldNames;
        public List<string> fieldTypeNames;
        public List<ETableFieldType> fieldTypes;
        public List<List<string>> value;

        public override string ToString()
        {
            var sb = new StringBuilder(1024);
            //sb.AppendLine("Defines: ");
            for (int i1 = 0; i1 < defines.Count; i1++)
            {
                sb.Append(defines[i1]).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Commits: ");
            for (int i1 = 0; i1 < fieldCommits.Count; i1++)
            {
                sb.Append(fieldCommits[i1]).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Types: ");
            for (int i1 = 0; i1 < fieldTypes.Count; i1++)
            {
                sb.Append(fieldTypes[i1].ToString()).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Names: ");
            for (int i1 = 0; i1 < fieldNames.Count; i1++)
            {
                sb.Append(fieldNames[i1]).Append("\t");
            }
            sb.AppendLine();
            //sb.AppendLine("Field Data: ");
            sb.AppendLine("Value Count: " + value.Count);
            for (int i1 = 0; i1 < value.Count; i1++)
            {
                var line = value[i1];
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
