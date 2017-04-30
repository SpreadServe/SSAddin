using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace SSAddin {
    // DataCache is the cache for non real time data like tiingo and quandl historical data result sets.
    // It also caches web socket data too. DataCache needs to be thread safe so worker threads can push
    // data in, and the Excel thread can pull data out. 
    class DataCache {
        protected Dictionary<string,List<String[]>> m_QCache = new Dictionary<string, List<String[]>>( );
        protected Dictionary<string, AnalyticDataPoint> m_GACache = new Dictionary<string, AnalyticDataPoint>();
        protected Dictionary<string, Dictionary<String, JObject>> m_BCache = new Dictionary<string, Dictionary<String, JObject>>( );
        protected Dictionary<string, List<SSTiingoHistPrice>> m_THPCache = new Dictionary<string, List<SSTiingoHistPrice>>( );
        protected Dictionary<string, string> m_WSCache = new Dictionary<string, string>( );
        protected static DataCache s_Instance;
        protected static object s_InstanceLock = new object( );
        protected static char[] s_delimiters = { '.' };

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

        public void UpdateTHPCache( string wkey, List<SSTiingoHistPrice> updates ) {
            lock (m_THPCache) {
                m_THPCache[wkey] = updates;
            }
        }

        public void UpdateGACache(string wkey, AnalyticDataPoint updates)
        {
            lock (m_GACache)
            {
                m_GACache[wkey] = updates;
            }
        }

        public void UpdateBareCache(string wkey, dynamic updates)
        {
            if ( updates == null) {
                Logr.Log( String.Format( "UpdateBareCache wkey({0}) updates==null!", wkey ) );
                return;
            }
            JToken metrics = null;
            if ( !updates.TryGetValue( "metrics", out metrics)) {
                Logr.Log( String.Format( "UpdateBareCache wkey({0}) couldn't get metrics!", wkey ) );
                return;
            }
            lock (m_BCache) {
                // the object returned by baremetrics always has an element
                // called 'metrics', and metrics is always an array of objects
                // that have data and human_date attributes. Beyond that we 
                // may have all sorts of other members and nesting.
                Dictionary<String, JObject> cached = null;
                if (!m_BCache.TryGetValue( wkey, out cached ))
                    cached = new Dictionary<string, JObject>( );
                foreach (JObject sum in metrics) {
                    JToken sdate = null;
                    if (sum.TryGetValue( "human_date", out sdate ))
                        cached[sdate.ToString( )] = sum;
                }
                m_BCache[wkey] = cached;
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

        public bool ContainsTiingoKey( string qkey ) {
            lock (m_THPCache) {
                return m_THPCache.ContainsKey( qkey );
            }
        }

        public bool ContainsGAnalyticsKey(string qkey)
        {
            lock (m_GACache) {
                return m_GACache.ContainsKey(qkey);
            }
        }

        public bool ContainsBareKey( string qkey ) {
            lock (m_BCache) {
                return m_BCache.ContainsKey( qkey );
            }
        }

        public void ClearTiingo( string qkey ) {
            lock (m_THPCache) {
                if (m_THPCache.ContainsKey( qkey ))
                    m_THPCache.Remove( qkey );
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

        public string GetTiingoCell( string qkey, int row, int col ) {
            lock (m_THPCache) {
                if (!m_THPCache.ContainsKey( qkey ))
                    return null;
                List<SSTiingoHistPrice> thplist = m_THPCache[qkey];
                if (row >= thplist.Count)
                    return null;
                SSTiingoHistPrice thp = thplist.ElementAt( row );
                PropertyInfo[] piarr = thp.GetType( ).GetProperties( );
                // String[] sarray = slist.ElementAt( row );
                if (col >= piarr.Length)
                    return null;
                PropertyInfo pi = piarr[col];
                return pi.GetValue( thp, BindingFlags.Default, null, null, null).ToString( );
            }
        }

        public string GetGAnalyticsCell(string qkey, int row, int col, bool hdrs)
        {
            lock (m_GACache) {
                if (!m_GACache.ContainsKey(qkey))
                    return null;
                if ( col < 0 || row < 0)
                    return null;
                AnalyticDataPoint adp = m_GACache[qkey];
                // If headers is true, we consider the headers as row 0, and the
                // first data record is row 1
                if (hdrs) {
                    if (row == 0) {
                        if (col >= adp.ColumnHeaders.Count)
                            return null;
                        return adp.ColumnHeaders[col].Name;
                    }
                    // adjust row down: row==1 refers to row 0 in adp.Rows
                    // when hdrs is true
                    row = row - 1;
                }
                if (row >= adp.Rows.Count)
                    return null;
                IList<string> record = adp.Rows[row];
                if (col >= record.Count)
                    return null;
                return record[col];
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

        public string GetBareField( string qkey, string sdate, string field) {
            // TODO: add more logging for failed nav thru the obj tree
            lock (m_BCache) {
                if (!m_BCache.ContainsKey( qkey ))
                    return null;
                Dictionary<String, JObject> cached = m_BCache[qkey];
                if (!cached.ContainsKey( sdate ))
                    return null;
                JObject jobj = cached[sdate];
                string[] subs = field.Split( s_delimiters );
                JToken jsub = null;
                JArray jarr = null;
                uint inx = 0;
                bool ok = false;
                foreach ( string sub in subs) {
                    if ( jobj != null) {
                        ok = jobj.TryGetValue( sub, out jsub );
                    }
                    else if ( jarr != null) {
                        ok = UInt32.TryParse( sub, out inx );
                        if (ok && inx < jarr.Count)
                            jsub = jarr[inx];
                        else
                            ok = false;
                    }
                    if (!ok)
                        return null;
                    // Have we hit an atomic, or do we go round again?
                    if ( jsub is JValue)
                        return jsub.ToString( );
                    // We haven't hit an atomic, so it's either an object or an array/
                    jobj = null;
                    jarr = null;
                    ok = false;
                    if (jsub is JObject)
                        jobj = jsub as JObject;
                    else if (jsub is JArray)
                        jarr = jsub as JArray;
                }
                return null;
            }
        }
    }
}
