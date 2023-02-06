namespace Saro.GTable
{
    internal static class Log
    {
        public static void Info(string value)
        {
#if UNITY_2017_1_OR_NEWER
            UnityEngine.Debug.Log(value);
#else
            System.Console.WriteLine(value);
#endif
        }

        internal static void Assert(bool v1, string v2)
        {
#if UNITY_2017_1_OR_NEWER
            UnityEngine.Debug.Assert(v1, v2);
#else
            System.Diagnostics.Debug.Assert(v1, v2);
#endif
        }
    }
}
