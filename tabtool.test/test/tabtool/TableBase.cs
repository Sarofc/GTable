using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace tabtool
{
    public abstract class TableBase<T, U>
    {
        public static U Get()
        {
            if (s_Instance == null)
            {
                s_Instance = Activator.CreateInstance<U>();
                (s_Instance as TableBase<T, U>).m_Datas = new Dictionary<int, T>();
            }

            return s_Instance;
        }

        protected static U s_Instance;

        protected Dictionary<int, T> m_Datas;

        public Dictionary<int, T> GetTable()
        {
            return m_Datas;
        }

        public T GetTableItem(int key)
        {
            T t;
            if (m_Datas.TryGetValue(key, out t))
            {
                return t;
            }
            return default(T);
        }

        public abstract bool Load();

        protected byte[] GetBytes(string tableName)
        {
            if (TableCfg.s_BytesLoader != null)
            {
                var path = Path.Combine(TableCfg.s_TableSrc, tableName);
                return TableCfg.s_BytesLoader(path);
            }
            return null;
        }

        public void Unload()
        {
            m_Datas.Clear();
            m_Datas = null;
            s_Instance = default;
        }

        
    }
}