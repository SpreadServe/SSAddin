using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SSAddin {
    // DataCache is the cache for non real time data like quandl historical data result sets.
    // DataCache needs to be thread safe so worker threads can push data in, and the Excel
    // thread can pull data out. 
    class DataCache {
        protected Dictionary<string,List<String[]>> m_QCache = new Dictionary<string, List<String[]>>( );
        protected Dictionary<string, string> m_WSCache = new Dictionary<string, string>( );
        protected static DataCache s_Instance;
        protected static object s_InstanceLock = new object( );

        private DataCache( ) {
        }

        public static DataCache Instance( ) {
            lock (s_InstanceLock) {
                if (s_Instance == null) {
                    s_Instance = new DataCache( );
                }
                return s_Instance;
            }
        }

        public int AddQuandlLine( string qkey, string[] line) {
            int count = 0;
            lock ( m_QCache) {
                if ( !m_QCache.ContainsKey( qkey))
                    m_QCache.Add( qkey, new List<string[]>( ));
                List<String[]> slist = m_QCache[qkey];
                slist.Add( line );
                count = slist.Count;
            }
            return count;
        }


        public void UpdateWSCache( string wkey, List<SSWebCell> updates) {
            lock (m_WSCache) {
                foreach (SSWebCell wc in updates) {
                    string key = String.Format( "{0}.{1}", wkey, wc.id);
                    m_WSCache[key] = wc.body;
                }
            }
        }

        public bool ContainsQuandlKey( string qkey ) {
            lock (m_QCache) {
                return m_QCache.ContainsKey( qkey );
            }
        }

        public void ClearQuandl( string qkey ) {
            lock (m_QCache) {
                if ( m_QCache.ContainsKey( qkey))
                    m_QCache.Remove( qkey );
            }
        }

        public string GetQuandlCell( string qkey, int row, int col ) {
            lock (m_QCache) {
                if (!m_QCache.ContainsKey( qkey ))
                    return null;
                List<String[]> slist = m_QCache[qkey];
                if (row >= slist.Count)
                    return null;
                String[] sarray = slist.ElementAt( row );
                if (col >= sarray.Length)
                    return null;
                return sarray[col];
            }
        }

        public string GetWSCell( string wkey, string ckey) {
            string key = String.Format( "{0}.{1}", wkey, ckey );
            lock (m_WSCache) {
                if (!m_WSCache.ContainsKey( key ))
                    return null;
                return m_WSCache[key];
            }
        }
    }
}
