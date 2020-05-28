//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.IO;
//using static tabtool.Program;

//namespace tabtool
//{
//    public class Item
//    {
//        public int id;
//        public string value;
//        public List<int> intList;
//    }

//    public class _CfgItemTable : TableBase<Item, _CfgItemTable>
//    {
//        public override bool Load()
//        {
//            var bytes = GetBytes("ITEM.txt");

//            using (var ms = new MemoryStream(bytes))
//            {
//                using (var br = new BinaryReader(ms))
//                {
//                    var version = br.ReadInt32();//version
//                    var typeCount = br.ReadInt32();
//                    var types = new byte[typeCount];
//                    for (int i = 0; i < typeCount; i++)
//                    {
//                        types[i] = br.ReadByte();
//                    }

//                    var dataLen = br.ReadInt32();
//                    for (int i = 0; i < dataLen; i++)
//                    {
//                        var data = new Item();
//                        data.id = br.ReadInt32();
//                        data.value = br.ReadString();

//                        var len = br.ReadInt32();
//                        data.intList = new List<int>(len);
//                        for (int j = 0; j < len; j++)
//                        {
//                            data.intList.Add(br.ReadInt32());
//                        }

//                        m_Datas[data.id] = data;
//                    }

//                    var sb = new StringBuilder();
//                    sb.Append("version: ").Append(version).AppendLine();
//                    foreach (var type in types)
//                    {
//                        sb.Append((ETableFieldType)type).Append("\t");
//                    }
//                    sb.AppendLine();
//                    foreach (var item in m_Datas.Values)
//                    {
//                        sb.Append(item.id).Append("\t").Append(item.value).
//                            AppendLine();
//                    }
//                    Console.WriteLine(sb.ToString());
//                }
//            }

//            return true;
//        }
//    }
//}
