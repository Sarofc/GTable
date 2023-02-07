#if UNITY_2017_1_OR_NEWER

namespace Saro.GTable
{
    public static partial class GTableConfig
    {
        public const string vfileName = "tables";

        public const string in_excel = "./Assets/Content/excel";

        public const string out_client = "./ExtraAssets/tables/data";

        public const string out_cs = "./Assets/Scripts/Hotfix/gen/Tables";
    }
}

#endif