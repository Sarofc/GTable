using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tabtool
{
    public class TableCfg
    {
        public static string s_TableSrc = "";
        public static Func<string, byte[]> s_BytesLoader;
    }
}
