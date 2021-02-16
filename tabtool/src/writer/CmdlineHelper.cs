using System;
using System.Linq;

namespace Saro.Table
{
    class CmdlineHelper
    {
        public CmdlineHelper(string[] args)
        {
            m_Args = args;
        }

        string[] m_Args;

        public bool Has(string s)
        {
            return m_Args.Count(p => p == s) > 0;
        }

        public string Get(string s)
        {
            for(int i = 0; i < m_Args.Count(); i++)
            {
                if (m_Args[i] == s && i + 1 < m_Args.Count())
                {
                    return m_Args[i + 1];
                }
            }
            return null;
        }
    }
}
