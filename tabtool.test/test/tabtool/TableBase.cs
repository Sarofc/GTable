using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace tabtool
{
    public abstract class TableBase<D, T>
    {
        public static T Get()
        {
            if (s_Instance == null)
            {
                s_Instance = Activator.CreateInstance<T>();
                (s_Instance as TableBase<D, T>).m_Datas = new Dictionary<int, D>();
            }

            return s_Instance;
        }

        protected static T s_Instance;

        protected Dictionary<int, D> m_Datas;

        public Dictionary<int, D> GetTable()
        {
            return m_Datas;
        }

        public D GetTableItem(int key)
        {
            if (m_Datas.TryGetValue(key, out D t))
            {
                return t;
            }
            return default;
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