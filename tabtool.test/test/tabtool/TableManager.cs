using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace tabtool
{
    public abstract class TableManager<T, U> : SingletonTable<U>
    {
       protected Dictionary<int, T> m_Datas = new Dictionary<int, T>();

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

        public void Unload()
        {
            m_Datas.Clear();
        }
    }
}