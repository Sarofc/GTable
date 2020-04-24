using System.Collections.Generic;

namespace tabtool
{
    enum ETableFieldType
    {
        Int,
        Float,
        String,
        Struct,
        IntList,
        FloatList,
        StringList,
        StructList,
    }

    class TableField
    {
        public ETableFieldType fieldType;
        public string fieldName;
        public string typeName;
        public string commits;

        public static string[] ts = { "int", "float", "string", "xxx", "List<int>", "List<float>", "List<string>", "xxx" };

        public string GetTypeNameOfStructList()
        {
            return typeName.Substring(0, typeName.Length - 1);
        }

        public string GetCsharpTypeName()
        {
            if (fieldType == ETableFieldType.Struct)
            {
                return typeName;
            }
            if (fieldType == ETableFieldType.StructList)
            {
                return string.Format("List<{0}>", typeName.Substring(0, typeName.Length - 1));
            }
            return ts[(int)fieldType];
        }
    }

    class TableMeta
    {
        public string TableName;
        public List<TableField> Fields = new List<TableField>();
        public string GetClassName()
        {
            return "Cfg" + TableName + "Table";
        }

        public string GetItemName()
        {
            return "Tbs" + TableName + "Item";
        }
    }
}
