﻿using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.CodeDom;
using Microsoft.CSharp;
using System.CodeDom.Compiler;

namespace Saro.Table
{
    class CodeGen
    {
        internal static void MakeCsharpFile(List<ExcelData> excelDatas, string codepath)
        {
            const string k_CsFileName = "CsvData.cs";
            string csfile = codepath + k_CsFileName;

            var sb = new StringBuilder(2048);

            var unit = new CodeCompileUnit();
            var tableNamespace = new CodeNamespace("Saro.Table");
            unit.Namespaces.Add(tableNamespace);
            tableNamespace.Imports.Add(new CodeNamespaceImport("System.Collections"));
            tableNamespace.Imports.Add(new CodeNamespaceImport("System.Collections.Generic"));
            tableNamespace.Imports.Add(new CodeNamespaceImport("System.IO"));
            tableNamespace.Imports.Add(new CodeNamespaceImport("System.Text"));

            var enumDefineList = new List<int>();
            for (int i = 0; i < excelDatas.Count; i++)
            {
                var meta = excelDatas[i];

                #region item class

                var itemClass = new CodeTypeDeclaration(meta.GetEntityClassName());
                itemClass.TypeAttributes = System.Reflection.TypeAttributes.Public | System.Reflection.TypeAttributes.Sealed;
                tableNamespace.Types.Add(itemClass);

                int index = -1;
                enumDefineList.Clear();
                foreach (var header in meta.header)
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

                var tableClass = new CodeTypeDeclaration(meta.GetWrapperClassName());
                tableClass.TypeAttributes = System.Reflection.TypeAttributes.Public | System.Reflection.TypeAttributes.Sealed;
                tableClass.BaseTypes.Add(new CodeTypeReference("BaseTable", new CodeTypeReference[] { new CodeTypeReference(meta.GetEntityClassName()), new CodeTypeReference(meta.GetWrapperClassName()) }));
                tableNamespace.Types.Add(tableClass);

                #region loadmethod

                var loadMethod = new CodeMemberMethod();
                tableClass.Members.Add(loadMethod);
                loadMethod.Name = "Load";
                loadMethod.ReturnType = new CodeTypeReference(typeof(bool));
                loadMethod.Attributes = MemberAttributes.Override | MemberAttributes.Public;

                sb.AppendLine("\t\t\tif (m_Loaded) return true;");
                sb.AppendLine($"\t\t\tvar bytes = GetBytes(\"{meta.tablName}.txt\");");
                sb.AppendLine();
                sb.AppendLine("\t\t\tusing (var ms = new MemoryStream(bytes))");
                sb.AppendLine("\t\t\t{");
                sb.AppendLine("\t\t\t\tusing (var br = new BinaryReader(ms))");
                sb.AppendLine("\t\t\t\t{");
                sb.AppendLine("\t\t\t\t\tvar version = br.ReadInt32();//version");
                sb.AppendLine("\t\t\t\t\tvar dataLen = br.ReadInt32();");
                sb.AppendLine("\t\t\t\t\tfor (int i = 0; i < dataLen; i++)");
                sb.AppendLine("\t\t\t\t\t{");
                sb.AppendLine($"\t\t\t\t\t\tvar data = new {meta.GetEntityClassName()}();");
                bool first = true;
                foreach (var header in meta.header)
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
                    else if (t == typeof(string))
                    {
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = br.ReadString();");

                    }
                    else if (t == typeof(List<byte>))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadUInt16();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadUInt16();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new List<byte>(len);");
                        sb.AppendLine($"\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}.Add(br.ReadByte());");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(List<int>))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadUInt16();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadUInt16();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new List<int>(len);");
                        sb.AppendLine($"\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}.Add(br.ReadInt32());");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(List<long>))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadUInt16();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadUInt16();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new List<long>(len);");
                        sb.AppendLine("\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}.Add(br.ReadInt64());");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(List<float>))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadUInt16();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadUInt16();");
                        }
                        sb.AppendLine($"\t\t\t\t\t\tdata.{header.fieldName} = new List<float>(len);");
                        sb.AppendLine("\t\t\t\t\t\tfor (int j = 0; j < len; j++)");
                        sb.AppendLine("\t\t\t\t\t\t{");
                        sb.AppendLine($"\t\t\t\t\t\t\tdata.{header.fieldName}.Add(br.ReadSingle());");
                        sb.AppendLine("\t\t\t\t\t\t}");
                    }
                    else if (t == typeof(Dictionary<int, int>))
                    {
                        if (first)
                        {
                            sb.AppendLine($"\t\t\t\t\t\tvar len = br.ReadUInt16();");
                            first = false;
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\t\tlen = br.ReadUInt16();");
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
                #endregion

                #region GetTableItem method

                var keyCount = meta.GetKeyCount();
                if (keyCount == 1)
                {
                    sb.AppendLine($"\t\tpublic static {meta.GetEntityClassName()} Query(int key1)");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine("\t\t\tvar key = KeyHelper.GetKey(key1);");
                }
                else if (keyCount == 2)
                {
                    sb.AppendLine($"\t\tpublic static {meta.GetEntityClassName()} Query(int key1, int key2)");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine("\t\t\tvar key = KeyHelper.GetKey(key1, key2);");
                }
                else if (keyCount == 3)
                {
                    sb.AppendLine($"\t\tpublic static {meta.GetEntityClassName()} Query(int key1, int key2, int key3)");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine("\t\t\tvar key = KeyHelper.GetKey(key1, key2, key3);");
                }
                else if (keyCount == 4)
                {
                    sb.AppendLine($"\t\tpublic static {meta.GetEntityClassName()} Query(int key1, int key2, int key3, int key4)");
                    sb.AppendLine("\t\t{");
                    sb.AppendLine("\t\t\tvar key = KeyHelper.GetKey(key1, key2, key3, key4);");
                }

                sb.AppendLine($"\t\t\tif (!Get().Load()) throw new System.Exception(\"load table failed.type: \" + nameof({meta.GetEntityClassName()}));");
                sb.AppendLine($"\t\t\tif (Get().m_Datas.TryGetValue(key, out {meta.GetEntityClassName()} t))");
                sb.AppendLine("\t\t\t{");
                sb.AppendLine("\t\t\t\treturn t;");
                sb.AppendLine("\t\t\t}");
                sb.AppendLine($"\t\t\tthrow new System.Exception(\"null table. type: \" + nameof({meta.GetEntityClassName()}));");
                sb.AppendLine("\t\t}");

                var queryMethod = new CodeSnippetTypeMember(sb.ToString());

                tableClass.Members.Add(queryMethod);

                sb.Clear();

                #endregion

                #region PrintTable method

                var printTableMethod = new CodeMemberMethod();
                tableClass.Members.Add(printTableMethod);
                printTableMethod.Name = "PrintTable";
                printTableMethod.ReturnType = new CodeTypeReference(typeof(string));
                printTableMethod.Attributes = MemberAttributes.Final | MemberAttributes.Public;

                sb.AppendLine("\t\t\tvar sb = new StringBuilder(1024);");
                sb.AppendLine("\t\t\tforeach (var data in m_Datas.Values)");
                sb.AppendLine("\t\t\t{");
                foreach (var header in meta.header)
                {
                    if (TableHelper.IgnoreHeader(header)) continue;

                    if (!TableHelper.s_TypeLut.TryGetValue(header.fieldTypeName, out Type t))
                    {
                        throw new Exception("type is not support: " + header.fieldTypeName);
                    }

                    if (t == typeof(byte)
                        || t == typeof(int)
                        || t == typeof(long)
                        || t == typeof(float)
                        || t == typeof(string))
                    {
                        sb.AppendLine($"\t\t\t\tsb.Append(data.{header.fieldName}).Append(\"\\t\");");
                    }
                    else if (t == typeof(List<byte>)
                            || t == typeof(List<int>)
                            || t == typeof(List<long>)
                            || t == typeof(List<float>))
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

                printTableMethod.Statements.Add(new CodeSnippetStatement(sb.ToString()));
                printTableMethod.Statements.Add(new CodeMethodReturnStatement(
                    new CodeSnippetExpression("sb.ToString()")));

                sb.Clear();
                #endregion

                #endregion

                #region enum define

                if (enumDefineList.Count > 0)
                {
                    var enumType = new CodeTypeDeclaration(meta.GetEnumName());
                    enumType.IsEnum = true;
                    tableNamespace.Types.Add(enumType);
                    for (int j = 0; j < meta.rowValues.Count; j++)
                    {
                        for (int k = 0; k < enumDefineList.Count; k++)
                        {
                            sb.Append(meta.rowValues[j][enumDefineList[k]]);
                            if (k != enumDefineList.Count - 1) sb.Append("_");
                        }
                        enumType.Members.Add(new CodeMemberField() { Name = sb.ToString(), InitExpression = new CodePrimitiveExpression(int.Parse(meta.rowValues[j][0])) });
                        sb.Clear();
                    }
                }

                #endregion
            }

            var tw = new IndentedTextWriter(new StreamWriter(csfile, false), "\t");
            {
                var provider = new CSharpCodeProvider();
                tw.WriteLine("//------------------------------------------------------------------------------");
                tw.WriteLine("// File   : {0}", k_CsFileName);
                tw.WriteLine("// Author : Saro");
                tw.WriteLine("// Time   : {0}", DateTime.Now.ToString());
                tw.WriteLine("//------------------------------------------------------------------------------");
                provider.GenerateCodeFromCompileUnit(unit, tw, new CodeGeneratorOptions() { BracingStyle = "C" });
                tw.Close();
            }
        }
    }
}
