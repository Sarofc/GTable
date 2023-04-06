#if UNITY_2017_1_OR_NEWER

using System;
using System.Threading.Tasks;
using UnityEditor;
using UnityEngine;
using UnityToolbarExtender;

namespace Saro.GTable
{
    [InitializeOnLoad]
    internal class UnityTableExporter
    {
        static class ToolbarStyles
        {
            public static readonly GUIStyle commandButtonStyle;
            public static Texture2D icon
            {
                get
                {
                    if (s_Icon == null)
                        s_Icon = AssetDatabase.LoadAssetAtPath<Texture2D>("Packages/com.saro.gtable/unity/icon.png");
                    return s_Icon;
                }
            }

            private static Texture2D s_Icon;

            static ToolbarStyles()
            {
                commandButtonStyle = new GUIStyle("Command")
                {
                    fontSize = 16,
                    alignment = TextAnchor.MiddleCenter,
                    imagePosition = ImagePosition.ImageOnly,
                    fontStyle = FontStyle.Bold,
                };
            }
        }

        static UnityTableExporter()
        {
            ToolbarExtender.LeftToolbarGUI.Add(OnToolbarGUI);
        }

        static void OnToolbarGUI()
        {
            GUILayout.FlexibleSpace();

            if (GUILayout.Button(new GUIContent("", ToolbarStyles.icon, "导出数据表"), ToolbarStyles.commandButtonStyle))
            {
                try
                {
                    ExportTableAsync();
                }
                catch (Exception e)
                {
                    Log.ERROR(e);
                }
            }
        }

        [MenuItem("GTable/Export Table")]
        public static async Task ExportTableAsync()
        {
            string exporterPath = $"--out_client {GTableConfig.out_client} --out_cs {GTableConfig.out_cs} --in_excel {GTableConfig.in_excel}";
            await TableExporter.ExportAsync(exporterPath.Split(" "));
        }
    }
}

#endif