using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.CodeDom;
using Microsoft.CSharp;
using System.CodeDom.Compiler;

namespace Saro.GTable
{
    class CodeGen
    {
        internal static void MakeCsharpFiles(IEnumerable<ExcelData> excelDatas, string codepath)
        {
            foreach (var excelData in excelDatas)
            {
                MakeCsharpFile(excelData, codepath);
            }
        }

        internal static void MakeCsharpFile(ExcelData excelData, string codepath)
        {
            //const string k_CsFileName = "CsvData.cs";
            //string csfile = codepath + k_CsFileName;

            var csfile = codepath;

            var sb = new StringBuilder(2048);

            var unit = new CodeCompileUnit();
            var tableNamespace = new CodeNamespace("Saro.GTable");
            unit.Namespaces.Add(tableNamespace);
            tableNamespace.Imports.Add(new CodeNamespaceImport("System.Collections"));
            tableNamespace.Imports.Add(new CodeNamespaceImport("System.Collections.Generic"));
            tableNamespace.Imports.Add(new CodeNamespaceImport("System.IO"));
            tableNamespace.Imports.Add(new CodeNamespaceImport("System.Text"));

            var enumDefineList = new List<int>();

            #region item class

            var itemClass = new CodeTypeDeclaration(excelData.GetEntityClassName());
            itemClass.TypeAttributes = System.Reflection.TypeAttributes.Public | System.Reflection.TypeAttributes.Sealed;
            tableNamespace.Types.Add(itemClass);

            int index = -1;
            enumDefineList.Clear();
            foreach (var header in excelData.header)
            {
                index++;
                if (header.define == TableHelper.HeaderFilter.k_ENUM_KEY)
                {
                    enumDefineList.Add(index);
                }
                if (TableHelper.IgnoreHeader(header)) continue;

                if (!TableHelper.s_TypeLut.TryGetValue(header.fieldTypeName, out Type t))
                {
                    throw new Exception("type is not support: " + header.fieldTypeName);
                }
                var memberFiled = new CodeMemberField(t, header.fieldName);
                memberFiled.Attributes = MemberAttributes.Public;

                memberFiled.Comments.Add(new CodeCommentStatement("<summary>", true));
                memberFiled.Comments.Add(new CodeCommentStatement(header.fieldComment, true));
                memberFiled.Comments.Add(new CodeCommentStatement("</summary>", true));
                itemClass.Members.Add(memberFiled);
            }

            #endregion

            #region table class

            var tableClass = new CodeTypeDeclaration(excelData.GetWrapperClassName());
            tableClass.TypeAttributes = System.Reflection.TypeAttributes.Public | System.Reflection.TypeAttributes.Sealed;
            tableClass.BaseTypes.Add(new CodeTypeReference("BaseTable", new CodeTypeReference[] { new CodeTypeReference(excelData.GetEntityClassName()), new CodeTypeReference(excelData.GetWrapperClassName()) }));
            tableNamespace.Types.Add(tableClass);

            #region load method
            {
                var loadMethod = new CodeMemberMethod();
                tableClass.Members.Add(loadMethod);
                loadMethod.Name = "Load";
                loadMethod.ReturnType = new CodeTypeReference(typeof(bool));
                loadMethod.Attributes = MemberAttributes.Override | MemberAttributes.Public;

                sb.AppendLine("\t\t\tif (m_Loaded) return true;");
                sb.AppendLine($"\t\t\tvar bytes = GetBytes(\"{excelData.tablName}.txt\");");
                sb.AppendLine();
                sb.AppendLine("\t\t\tusing (var ms = new MemoryStream(bytes, false))");
                sb.AppendLine("\t\t\t{");
                sb.AppendLine("\t\t\t\tusing (var br = new BinaryReader(ms))");
                sb.AppendLine("\t\t\t\t{");
                sb.AppendLine("\t\t\t\t\tvar version = br.ReadInt32();//version");
                sb.AppendLine("\t\t\t\t\tif (version != TableLoader.k_DataVersion)\n\t\t\t\t\t\tthrow new System.Exception($\"table error version. file:{version}  exe:{TableLoader.k_DataVersion}\");\n");
                sb.AppendLine("\t\t\t\t\tvar dataLen = br.ReadInt32();");
                sb.AppendLine("\t\t\t\t\tfor (int i = 0; i < dataLen; i++)");
                sb.AppendLine("\t\t\t\t\t{");
                sb.AppendLine($"\t\t\t\t\t\tvar data = new {excelData.GetEntityClassName()}();");
                bool first = true;
                foreach (var header in excelData.header)
                {
                    if (TableHelper.IgnoreHeader(header)) continue;

                    if (!TableHelper.s_TypeLut.TryGetValue(header.fieldTypeName, out Type t))
                    {
                        throw new Exception("type is not support: " + header.fieldTypeName);
                    }
                    if (t == typeof(byte))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadByte();");
                    }
                    else if (t == typeof(int))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadInt32();");
                    }
                    else if (t == typeof(long))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadInt64();");
                    }
                    else if (t == typeof(float))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadSingle();");
                    }
                    else if (t == typeof(bool))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadBoolean();");
                    }
                    else if (t == typeof(string))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadString();");

                    }
                    else if (t == typeof(byte[]))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new byte[len];");
                        sb.AppendLine($"\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}[j] = br.ReadByte();");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(int[]))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new int[len];");
                        sb.AppendLine($"\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}[j] = br.ReadInt32();");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(long[]))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new long[len];");
                        sb.AppendLine("\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}[j] = br.ReadInt64();");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(float[]))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new float[len];");
                        sb.AppendLine("\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}[j] = br.ReadSingle();");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(Dictionary<int, int>))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new Dictionary<int, int>(len);");
                        sb.AppendLine("\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tvar key = br.ReadInt32();");
                        sb.AppendLine($"\t\t\t\t\t\t\tvar val = br.ReadInt32();");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}.Add(key, val);");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                }
                sb.AppendLine("\t\t\t\t\t\tvar _key = br.ReadUInt64();");
                sb.AppendLine("\t\t\t\t\t\tm_Datas[_key] = data;");
                sb.AppendLine("\t\t\t\t\t}");
                sb.AppendLine("\t\t\t\t}");
                sb.AppendLine("\t\t\t}");
                sb.AppendLine("\t\t\tm_Loaded = true;");
                loadMethod.Statements.Add(new CodeSnippetStatement(sb.ToString()));
                loadMethod.Statements.Add(new CodeMethodReturnStatement(
                    new CodeSnippetExpression("true")));

                sb.Clear();
            }
            #endregion

            #region loadasync method
            {
                var loadAsyncMethod = new CodeMemberMethod();
                tableClass.Members.Add(loadAsyncMethod);
                loadAsyncMethod.Name = "LoadAsync";
                loadAsyncMethod.ReturnType = new CodeTypeReference("async System.Threading.Tasks.ValueTask<bool>");
                loadAsyncMethod.Attributes = MemberAttributes.Override | MemberAttributes.Public;

                sb.AppendLine("\t\t\tif (m_Loaded) return true;");
                sb.AppendLine($"\t\t\tvar bytes = await GetBytesAsync(\"{excelData.tablName}.txt\");");
                sb.AppendLine();
                sb.AppendLine("\t\t\tusing (var ms = new MemoryStream(bytes, false))");
                sb.AppendLine("\t\t\t{");
                sb.AppendLine("\t\t\t\tusing (var br = new BinaryReader(ms))");
                sb.AppendLine("\t\t\t\t{");
                sb.AppendLine("\t\t\t\t\tvar version = br.ReadInt32();//version");
                sb.AppendLine("\t\t\t\t\tif (version != TableLoader.k_DataVersion)\n\t\t\t\t\t\tthrow new System.Exception($\"table error version. file:{version}  exe:{TableLoader.k_DataVersion}\");\n");
                sb.AppendLine("\t\t\t\t\tvar dataLen = br.ReadInt32();");
                sb.AppendLine("\t\t\t\t\tfor (int i = 0; i < dataLen; i++)");
                sb.AppendLine("\t\t\t\t\t{");
                sb.AppendLine($"\t\t\t\t\t\tvar data = new {excelData.GetEntityClassName()}();");
                bool first = true;
                foreach (var header in excelData.header)
                {
                    if (TableHelper.IgnoreHeader(header)) continue;

                    if (!TableHelper.s_TypeLut.TryGetValue(header.fieldTypeName, out Type t))
                    {
                        throw new Exception("type is not support: " + header.fieldTypeName);
                    }
                    if (t == typeof(byte))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadByte();");
                    }
                    else if (t == typeof(int))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadInt32();");
                    }
                    else if (t == typeof(long))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadInt64();");
                    }
                    else if (t == typeof(float))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadSingle();");
                    }
                    else if (t == typeof(bool))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadBoolean();");
                    }
                    else if (t == typeof(string))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadString();");

                    }
                    else if (t == typeof(byte[]))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new byte[len];");
                        sb.AppendLine($"\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}[j] = br.ReadByte();");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(int[]))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new int[len];");
                        sb.AppendLine($"\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}[j] = br.ReadInt32();");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(long[]))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new long[len];");
                        sb.AppendLine("\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}[j] = br.ReadInt64();");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(float[]))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new float[len];");
                        sb.AppendLine("\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}[j] = br.ReadSingle();");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(Dictionary<int, int>))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadByte();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadByte();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new Dictionary<int, int>(len);");
                        sb.AppendLine("\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tvar key = br.ReadInt32();");
                        sb.AppendLine($"\t\t\t\t\t\t\tvar val = br.ReadInt32();");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}.Add(key, val);");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                }
                sb.AppendLine("\t\t\t\t\t\tvar _key = br.ReadUInt64();");
                sb.AppendLine("\t\t\t\t\t\tm_Datas[_key] = data;");
                sb.AppendLine("\t\t\t\t\t}");
                sb.AppendLine("\t\t\t\t}");
                sb.AppendLine("\t\t\t}");
                sb.AppendLine("\t\t\tm_Loaded = true;");
                loadAsyncMethod.Statements.Add(new CodeSnippetStatement(sb.ToString()));
                loadAsyncMethod.Statements.Add(new CodeMethodReturnStatement(
                    new CodeSnippetExpression("true")));

                sb.Clear();
            }
            #endregion

            #region query method
            {
                var keyCount = excelData.GetKeyCount();
                var keyNames = excelData.GetKeyNames();
                var combinedKeyName = "__combinedkey";
                if (keyCount == 1)
                {
                    sb.AppendLine($"\t\tpublic static {excelData.GetEntityClassName()} Query(int {keyNames[0]})");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine($"\t\t\tvar {combinedKeyName} = KeyHelper.GetKey({keyNames[0]});");
                }
                else if (keyCount == 2)
                {
                    sb.AppendLine($"\t\tpublic static {excelData.GetEntityClassName()} Query(int {keyNames[0]}, int {keyNames[1]})");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine($"\t\t\tvar {combinedKeyName} = KeyHelper.GetKey({keyNames[0]}, {keyNames[1]});");
                }
                else if (keyCount == 3)
                {
                    sb.AppendLine($"\t\tpublic static {excelData.GetEntityClassName()} Query(int {keyNames[0]}, int {keyNames[1]}, int {keyNames[2]})");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine($"\t\t\tvar {combinedKeyName} = KeyHelper.GetKey({keyNames[0]}, {keyNames[1]}, {keyNames[2]});");
                }
                else if (keyCount == 4)
                {
                    sb.AppendLine($"\t\tpublic static {excelData.GetEntityClassName()} Query(int {keyNames[0]}, int {keyNames[1]}, int {keyNames[2]}, int {keyNames[3]})");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine($"\t\t\tvar {combinedKeyName} = KeyHelper.GetKey({keyNames[0]}, {keyNames[1]}, {keyNames[2]}, {keyNames[3]});");
                }

                sb.AppendLine($"\t\t\tif (!Get().Load()) throw new System.Exception(\"load table failed.type: \" + nameof({excelData.GetEntityClassName()}));");
                sb.AppendLine($"\t\t\tif (Get().m_Datas.TryGetValue({combinedKeyName}, out {excelData.GetEntityClassName()} t))");
                sb.AppendLine("\t\t\t{");
                sb.AppendLine("\t\t\t\treturn t;");
                sb.AppendLine("\t\t\t}");
                sb.AppendLine($"\t\t\tthrow new System.Exception(\"null table. type: \" + nameof({excelData.GetEntityClassName()}));");
                sb.AppendLine("\t\t}");

                var queryMethod = new CodeSnippetTypeMember(sb.ToString());

                tableClass.Members.Add(queryMethod);

                sb.Clear();
            }
            #endregion


            #region queryasync method
            {
                var keyCount = excelData.GetKeyCount();
                var keyNames = excelData.GetKeyNames();
                var combinedKeyName = "__combinedkey";
                if (keyCount == 1)
                {
                    sb.AppendLine($"\t\tpublic static async System.Threading.Tasks.ValueTask<{excelData.GetEntityClassName()}> QueryAsync(int {keyNames[0]})");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine($"\t\t\tvar {combinedKeyName} = KeyHelper.GetKey({keyNames[0]});");
                }
                else if (keyCount == 2)
                {
                    sb.AppendLine($"\t\tpublic static async System.Threading.Tasks.ValueTask<{excelData.GetEntityClassName()}> QueryAsync(int {keyNames[0]}, int {keyNames[1]})");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine($"\t\t\tvar {combinedKeyName} = KeyHelper.GetKey({keyNames[0]}, {keyNames[1]});");
                }
                else if (keyCount == 3)
                {
                    sb.AppendLine($"\t\tpublic static async System.Threading.Tasks.ValueTask<{excelData.GetEntityClassName()}> QueryAsync(int {keyNames[0]}, int {keyNames[1]}, int {keyNames[2]})");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine($"\t\t\tvar {combinedKeyName} = KeyHelper.GetKey({keyNames[0]}, {keyNames[1]}, {keyNames[2]});");
                }
                else if (keyCount == 4)
                {
                    sb.AppendLine($"\t\tpublic static async System.Threading.Tasks.ValueTask<{excelData.GetEntityClassName()}> QueryAsync(int {keyNames[0]}, int {keyNames[1]}, int {keyNames[2]}, int {keyNames[3]})");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine($"\t\t\tvar {combinedKeyName} = KeyHelper.GetKey({keyNames[0]}, {keyNames[1]}, {keyNames[2]}, {keyNames[3]});");
                }

                sb.AppendLine($"\t\t\tvar result = await Get().LoadAsync();");
                sb.AppendLine($"\t\t\tif (!result) throw new System.Exception(\"load table failed.type: \" + nameof({excelData.GetEntityClassName()}));");
                sb.AppendLine($"\t\t\tif (Get().m_Datas.TryGetValue({combinedKeyName}, out {excelData.GetEntityClassName()} t))");
                sb.AppendLine("\t\t\t{");
                sb.AppendLine("\t\t\t\treturn t;");
                sb.AppendLine("\t\t\t}");
                sb.AppendLine($"\t\t\tthrow new System.Exception(\"null table. type: \" + nameof({excelData.GetEntityClassName()}));");
                sb.AppendLine("\t\t}");

                var queryMethod = new CodeSnippetTypeMember(sb.ToString());

                tableClass.Members.Add(queryMethod);

                sb.Clear();
            }
            #endregion


            #region PrintTable method

            var printTableMethod = new CodeMemberMethod();
            tableClass.Members.Add(printTableMethod);
            printTableMethod.Name = "PrintTable";
            printTableMethod.ReturnType = new CodeTypeReference(typeof(string));
            printTableMethod.Attributes = MemberAttributes.Final | MemberAttributes.Public;

            sb.AppendLine("\t\t\tStringBuilder sb = null;");

            sb.AppendLine("#if ENABLE_TABLE_LOG");

            sb.AppendLine("\t\t\tsb = new StringBuilder(2048);");

            sb.AppendLine("\t\t\tforeach (var data in m_Datas.Values)");
            sb.AppendLine("\t\t\t{");
            foreach (var header in excelData.header)
            {
                if (TableHelper.IgnoreHeader(header)) continue;

                if (!TableHelper.s_TypeLut.TryGetValue(header.fieldTypeName, out Type t))
                {
                    throw new Exception("type is not support: " + header.fieldTypeName);
                }

                if (t.IsValueType
                    || t == typeof(string))
                {
                    sb.AppendLine($"\t\t\t\tsb.Append(data.{header.fieldName}).Append(\"\\t\");");
                }
                else if (t == typeof(byte[])
                        || t == typeof(int[])
                        || t == typeof(long[])
                        || t == typeof(float[]))
                {
                    sb.AppendLine($"\t\t\t\tsb.Append(string.Join(\",\", data.{header.fieldName})).Append(\"\\t\");");
                }
                else if (t == typeof(Dictionary<int, int>))
                {
                    sb.AppendLine($"\t\t\t\tsb.Append(string.Join(\",\", data.{header.fieldName})).Append(\"\\t\");");
                }
            }
            sb.AppendLine("\t\t\t\tsb.AppendLine();");
            sb.AppendLine("\t\t\t}");

            sb.AppendLine("#endif");

            printTableMethod.Statements.Add(new CodeSnippetStatement(sb.ToString()));
            printTableMethod.Statements.Add(new CodeMethodReturnStatement(new CodeSnippetExpression("sb?.ToString()")));

            sb.Clear();

            #endregion

            #endregion

            #region enum define

            if (enumDefineList.Count > 0)
            {
                var enumType = new CodeTypeDeclaration(excelData.GetEnumName());
                enumType.IsEnum = true;
                tableNamespace.Types.Add(enumType);
                for (int j = 0; j < excelData.rowValues.Count; j++)
                {
                    for (int k = 0; k < enumDefineList.Count; k++)
                    {
                        sb.Append(excelData.rowValues[j][enumDefineList[k]]);
                        if (k != enumDefineList.Count - 1) sb.Append("_");
                    }
                    enumType.Members.Add(new CodeMemberField() { Name = sb.ToString(), InitExpression = new CodePrimitiveExpression(int.Parse(excelData.rowValues[j][0])) });
                    sb.Clear();
                }
            }

            #endregion

#if NET6_OR_NEWER && false
            var options = new FileStreamOptions { Mode = FileMode.Create, Access = FileAccess.Write, Share = FileShare.ReadWrite };
            using var tw = new IndentedTextWriter(new StreamWriter(csfile, options), "\t");
#else
            using var fs = new FileStream(csfile, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
            //using var tw = new IndentedTextWriter(new StreamWriter(fs), "\t"); // bug? 文本不全
            using var tw = new StreamWriter(fs);
#endif
            var provider = new CSharpCodeProvider();
            tw.WriteLine("//------------------------------------------------------------------------------");
            tw.WriteLine("// File   : {0}", Path.GetFileName(codepath));
            tw.WriteLine("// Author : Saro");
            tw.WriteLine("// Time   : {0}", DateTime.Now.ToString());
            tw.WriteLine("//------------------------------------------------------------------------------");
            provider.GenerateCodeFromCompileUnit(unit, tw, new CodeGeneratorOptions() { BracingStyle = "C" });
        }
    }
}
