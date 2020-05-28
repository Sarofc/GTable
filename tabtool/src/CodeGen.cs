using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.CodeDom;
using Microsoft.CSharp;
using System.CodeDom.Compiler;

namespace tabtool
{
    class CodeGen
    {
        internal static void MakeCsharpFile(List<ExcelData> excelDatas, string codepath)
        {
            const string k_CsFileName = "CfgData.cs";
            string csfile = codepath + k_CsFileName;

            var sb = new StringBuilder(2048);

            var unit = new CodeCompileUnit();
            var tableNamespace = new CodeNamespace("tabtool");
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

                var itemClass = new CodeTypeDeclaration(meta.GetItemName());
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

                    var memberFiled = new CodeMemberField(TableHelper.s_TypeLut[header.fieldTypeName], header.fieldName);
                    memberFiled.Attributes = MemberAttributes.Public;

                    memberFiled.Comments.Add(new CodeCommentStatement("<summary>", true));
                    memberFiled.Comments.Add(new CodeCommentStatement(header.fieldComment, true));
                    memberFiled.Comments.Add(new CodeCommentStatement("</summary>", true));
                    itemClass.Members.Add(memberFiled);
                }

                #endregion

                #region table class

                var tableClass = new CodeTypeDeclaration(meta.GetClassName());
                tableClass.BaseTypes.Add(new CodeTypeReference("TableBase", new CodeTypeReference[] { new CodeTypeReference(meta.GetItemName()), new CodeTypeReference(meta.GetClassName()) }));
                tableNamespace.Types.Add(tableClass);

                #region loadmethod

                var loadMethod = new CodeMemberMethod();
                tableClass.Members.Add(loadMethod);
                loadMethod.Name = "Load";
                loadMethod.ReturnType = new CodeTypeReference(typeof(bool));
                loadMethod.Attributes = MemberAttributes.Override | MemberAttributes.Public;

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
                sb.AppendLine($"\t\t\t\t\t\tvar data = new {meta.GetItemName()}();");
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
                }
                sb.AppendLine($"\t\t\t\t\t\tm_Datas[data.{meta.header[0].fieldName}] = data;");
                sb.AppendLine("\t\t\t\t\t}");
                sb.AppendLine("\t\t\t\t}");
                sb.AppendLine("\t\t\t}");
                loadMethod.Statements.Add(new CodeSnippetStatement(sb.ToString()));
                loadMethod.Statements.Add(new CodeMethodReturnStatement(
                    new CodeSnippetExpression("true")));

                sb.Clear();
                #endregion

                #region tostring method

                var toStringMethod = new CodeMemberMethod();
                tableClass.Members.Add(toStringMethod);
                toStringMethod.Name = "ToString";
                toStringMethod.ReturnType = new CodeTypeReference(typeof(string));
                toStringMethod.Attributes = MemberAttributes.Override | MemberAttributes.Public;

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
                }
                sb.AppendLine("\t\t\t\tsb.AppendLine();");
                sb.AppendLine("\t\t\t}");

                toStringMethod.Statements.Add(new CodeSnippetStatement(sb.ToString()));
                toStringMethod.Statements.Add(new CodeMethodReturnStatement(
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
