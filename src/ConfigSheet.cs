using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

// https://www.quandl.com/api/v1/datasets/WIKI/AAPL.csv
// https://www.quandl.com/api/v1/datasets/WIKI/AAPL.json?trim_start=1985-05-01&trim_end=1997-07-01&sort_order=asc&column=4&collapse=quarterly&transformation=rdiff


namespace SSAddin {

    // Wrapper for all the config info in the s2cfg sheet, if it exists

    class ConfigSheet {

        protected static String s_QuandlBaseURL = "https://www.quandl.com/api/v1/datasets";
        protected static String s_TiingoBaseURL = "https://api.tiingo.com/tiingo";
        protected static String s_BareBaseURL = "https://api.baremetrics.com/v1/metrics";
        protected static String s_BareSandBaseURL = "https://api-sandbox.baremetrics.com/v1/metrics";
        protected static Dictionary<string, Func<object, string>> s_QuandlQueryFieldConverters = new Dictionary<string, Func<object, string>>( );
        protected static Dictionary<string, Func<object, string>> s_TiingoQueryFieldConverters = new Dictionary<string, Func<object, string>>( );
        protected static Dictionary<string, Func<object, string>> s_BareQueryFieldConverters = new Dictionary<string, Func<object, string>>();
        protected static string[] s_ProxyKeys = { "http_proxy_host", "http_proxy_port", "http_proxy_user", "http_proxy_password" };

        protected Dictionary<string, string> m_ConfigCache = new Dictionary<string, string>();

        public ConfigSheet( ) {
            // Quandl and Tiingo funcs expect Excel dates. In the GUI Excel recognises the date, and converts to a number
            s_QuandlQueryFieldConverters["trim_start"] = ExcelDateNumberToString;
            s_QuandlQueryFieldConverters["trim_end"] = ExcelDateNumberToString;
            s_TiingoQueryFieldConverters["startDate"] = ExcelDateNumberToString;
            s_TiingoQueryFieldConverters["endDate"] = ExcelDateNumberToString;
            // For baremetrics queries date params should be supplied by s2today, which arrives in here as
            // a string, not a number like an Excel date.
            // s_BareQueryFieldConverters["start_date"] = ExcelDateNumberToString;
            // s_BareQueryFieldConverters["end_date"] = ExcelDateNumberToString;
        }

        public object GetCell( int row, int col ) {
            ExcelReference xlref = new ExcelReference( row, row, col, col, "s2cfg" );
            return xlref.GetValue( );
        }

        public String GetCellAsString( int row, int col ) {
            object val = GetCell( row, col);
            if ( val == ExcelEmpty.Value)
                return "";
            return val.ToString( );
        }

        public string ExcelDateNumberToString( object dn ) {
            DateTime dt = DateTime.FromOADate( Convert.ToDouble( dn));
            // "u" format is 2008-06-15 21:15:07Z
            string sdt = dt.ToString( "u" );
            // Throw away the time and just keep the date...
            return sdt.Substring( 0, 10 );
        }

        public String BuildQuandlQuery( Dictionary<string, string> qterms ) {
            if ( !qterms.ContainsKey( "dataset"))
                return "";
            StringBuilder sb = new StringBuilder( s_QuandlBaseURL);
            sb.Append( String.Format( "/{0}.csv", qterms["dataset"]));
            qterms.Remove( "dataset");
            string prefix = "?";
            string auth_token = GetQueryConfig( "quandl", "auth_token" );
            if (auth_token != "") {
                sb.Append( String.Format( "{0}{1}={2}", prefix, "auth_token", auth_token ) );
                prefix = "&";
            }
            string val;
            foreach ( KeyValuePair<string, string> item in qterms) {
                if (s_QuandlQueryFieldConverters.ContainsKey( item.Key )) {
                    Func<object, string> converter = s_QuandlQueryFieldConverters[item.Key];
                    val = converter( item.Value );
                }
                else {
                    val = item.Value.ToString( );
                }
                sb.Append( String.Format( "{0}{1}={2}", prefix, item.Key, val));
                prefix = "&";
            }
            return sb.ToString( );
        }

        public String BuildTiingoQuery( Dictionary<string, string> qterms ) {
            // Must specify a ticker symbol and a root (daily|funds)
            if (!qterms.ContainsKey( "ticker" ) || !qterms.ContainsKey("root"))
                return "";
            // Now deal with the minimal essential we need to build a valid tiingo query
            StringBuilder sb = new StringBuilder( s_TiingoBaseURL );
            sb.Append( String.Format( "/{0}/{1}", qterms["root"], qterms["ticker"] ) );
            qterms.Remove( "ticker" );
            qterms.Remove( "root" );
            if (qterms.ContainsKey( "leaf" )) {
                sb.Append( String.Format( "/{0}", qterms["leaf"] ) );
                qterms.Remove( "leaf" );
            }
            // Are there any more values in the Dict? Maybe startDate and endDate...
            if (qterms.Count > 0) {
                string prefix = "?";
                string val;
                foreach (KeyValuePair<string, string> item in qterms) {
                    if (s_TiingoQueryFieldConverters.ContainsKey( item.Key )) {
                        Func<object, string> converter = s_TiingoQueryFieldConverters[item.Key];
                        val = converter( item.Value );
                    }
                    else {
                        val = item.Value.ToString( );
                    }
                    sb.Append( String.Format( "{0}{1}={2}", prefix, item.Key, val ) );
                    prefix = "&";
                }
            }
            return sb.ToString( );
        }

        public String BuildBareQuery(Dictionary<string, string> qterms)
        {
            // Must specify a query type
            if (!qterms.ContainsKey("qtype"))
                return "";
            // Now deal with the minimal essential we need to build a valid baremetrics query
            bool sandbox = false;
            if (qterms.ContainsKey("sandbox")) {
                try {
                    sandbox = Convert.ToBoolean(qterms["sandbox"]);
                }
                catch ( Exception ex) {
                    Logr.Log(String.Format("BuildBareQuery: bad sandbox value {0}\n{1}", qterms["sandbox"].ToString( ), ex));
                }
                qterms.Remove( "sandbox" );
            }
            StringBuilder sb = new StringBuilder(sandbox ? s_BareSandBaseURL : s_BareBaseURL);
            String qtype = qterms["qtype"].ToString( );
            if (qtype == "summary") {
                // null op for summary as we append the start and end dates below
            }
            else if (qtype == "metric") {
                if ( qterms.ContainsKey( "metric")) {
                    sb.Append( String.Format( "/{0}", qterms["metric"]));
                    qterms.Remove( "metric" );
                }
                else {
                    Logr.Log(String.Format("BuildBareQuery: qtype==metric, but no metric specified eg metric=mrr"));
                    return "";
                }
            }
            else if (qtype == "plan") {
                if (qterms.ContainsKey( "metric" )) {
                    sb.Append( String.Format( "/{0}/plans", qterms["metric"] ) );
                    qterms.Remove( "metric" );
                }
                else {
                    Logr.Log( String.Format( "BuildBareQuery: qtype==plan, but no metric specified eg metric=mrr" ) );
                    return "";
                }
            }
            else {
                return "";
            }
            qterms.Remove("qtype");
            // Are there any more values in the Dict? Maybe start_date and end_date...
            if (qterms.Count > 0)
            {
                string prefix = "?";
                string val;
                foreach (KeyValuePair<string, string> item in qterms)
                {
                    if (s_BareQueryFieldConverters.ContainsKey(item.Key))
                    {
                        Func<object, string> converter = s_BareQueryFieldConverters[item.Key];
                        val = converter(item.Value);
                    }
                    else
                    {
                        val = item.Value.ToString();
                    }
                    sb.Append(String.Format("{0}{1}={2}", prefix, item.Key, val));
                    prefix = "&";
                }
            }
            return sb.ToString();
        }

        public int FindRow( string c0, string c1, string c2, int startRow = 0)
        {
            // We're looking for a row that has c0 in the first cell, c1 in the second,
            // and then c2 in the third.
            int row = startRow;
            string a, b, c;
            do {    // keep going as long as the first field in a row isn't empty
                a = GetCellAsString( row, 0 );
                b = GetCellAsString( row, 1 );
                c = GetCellAsString( row, 2 );
                if ( a == c0 && b == c1 && c == c2)
                    return row;
                row++;
            } while (a != null && a != "");
            return -1;
        }

        public int FindRow(string c0, string c1, string c2, string c3)
        {
            // We're looking for a row that has c0 in the first cell, c1 in the second,
            // c2 in the third and c3 in the fourth
            string d;
            int row = FindRow(c0, c1, c2, 0);
            while ( row != -1) {
                d = GetCellAsString(row, 3);
                if (d == c3)
                    return row;
                row = FindRow(c0, c1, c2, row + 1);
            }
            return row;
        }

        public String GetQueryURL( String qtype, String qkey) {
            Dictionary<string, string> qterms = GetTerms( qtype, "query", qkey );
            if ( qtype == "quandl")
                return BuildQuandlQuery( qterms );
            if (qtype == "baremetrics")
                return BuildBareQuery(qterms);
            return BuildTiingoQuery( qterms );
        }

        public Dictionary<string, string> GetTerms( String qtype, String col2, String qkey ) {
            // We're looking for a row that has 'quandl', 'tiingo', 'baremetrics' or ganalytics in the first cell,
            // query in the second, and then qkey in the third.
            int row = FindRow( qtype, col2, qkey );
            if (row == -1) {
                Logr.Log( String.Format( "GetTerms: couldn't find {0} {1} {2}", qtype, col2, qkey ) );
                return null;
            }
            return GetKeyValPairs( row);
        }

        protected Dictionary<string, string> GetKeyValPairs( int row)
        {
            int col = 3;
            string name;
            string val;
            Dictionary<string, string> qterms = new Dictionary<string, string>();
            do {
                name = GetCellAsString(row, col);
                val = GetCellAsString(row, col + 1);
                if (name != null && name != "")
                    qterms.Add(name, val);
                col += 2;
            } while (name != null && name != "");
            return qterms;
        }

        public String GetQueryConfig(String qtype, String ckey)
        {
            // Get config for all queries of type qtype. For instance,
            // the auth_key for all quandl queries.
            string xkey = String.Format("{0}.{1}", qtype, ckey);
            if (m_ConfigCache.ContainsKey(xkey))
            {
                return m_ConfigCache[xkey];
            }
            // Look for config in .Net .config file too. Could be auth_token or http_proxy,
            // because we don't want to expose it in the sheet, and we don't want to repeat
            // in info in every sheet.
            string token = ConfigurationManager.AppSettings.Get( xkey );
            if (token != null && token != "") {
                m_ConfigCache[xkey] = token;
                Logr.Log( String.Format( "GetQueryConfig: using ssaddin.xll.config for {0}:{1}", xkey, token ) );
                return token;
            }
            // We're looking for a row that has qtype [quandl|tiingo] in the first cell, config in the second,
            // and then ckey in the third.
            int row = FindRow( qtype, "config", ckey );
            if (row == -1) {
                Logr.Log( String.Format( "GetQueryConfig: couldn't find {0}.{1}", qtype, ckey ) );
                // Caching "" as a value will ensure we get this logged once only
                m_ConfigCache[xkey] = "";
                return "";
            }
            string val = GetCellAsString( row, 3 );
            Logr.Log( String.Format( "GetQueryConfig: returning row {0} col 3 value:{1}", row, val ) );
            m_ConfigCache[xkey] = val;
            return val;
        }

        public int GetQueryConfigAsInt(String qtype, String qkey, String ckey)
        {
            // Get config for a specific query qkey, rather than all queries of type qtype.
            // For example, get xoffset (ckey) for a ganalytics (qtype) query called
            // metrics1 (qkey).
            int rv = 0;
            string xkey = String.Format("{0}.{1}.{2}", qtype, qkey, ckey);
            if (m_ConfigCache.ContainsKey(xkey)) {
                if (Int32.TryParse(m_ConfigCache[xkey], out rv))
                    return rv;
            }
            // Look for config in .Net .config file too. Could be auth_token or http_proxy,
            // because we don't want to expose it in the sheet, and we don't want to repeat
            // in info in every sheet.
            string token = ConfigurationManager.AppSettings.Get(xkey);
            if (token != null && token != "") {
                m_ConfigCache[xkey] = token;
                if (Int32.TryParse(token, out rv))
                    Logr.Log(String.Format("GetQueryConfigAsInt: using ssaddin.xll.config for {0}:{1}", xkey, token));
                else 
                    Logr.Log(String.Format("GetQueryConfigAsInt: INT ERROR ssaddin.xll.config for {0}:{1}", xkey, token));
                return rv;
            }
            // We're looking for a row that has qtype [quandl|tiingo] in the first cell, config in the second,
            // and then qkey in the third.
            int row = FindRow(qtype, "config", qkey, ckey);
            if (row != -1) {
                string val = GetCellAsString(row, 4);
                m_ConfigCache[xkey] = val;
                if (Int32.TryParse(val, out rv)) {
                    Logr.Log(String.Format("GetQueryConfigAsInt: returning row {0} col 4 value:{1}", row, val));
                    return rv;
                }
                Logr.Log(String.Format("GetQueryConfigAsInt: INT ERROR row {0} col 3 value:{1}", row, val));
                // Caching "" as a value will ensure we get this logged once only
                m_ConfigCache[xkey] = "";
            }
            return rv;
        }

        public Tuple<String,DateTime?,DateTime?> GetCronTab( String ctabkey ) {
            // We're looking for a row that has 'cron' in the first cell, tab in the second,
            // and then ctabkey in the third.
            int row = FindRow( "cron", "tab", ctabkey );
            if (row == -1) {
                Logr.Log( String.Format( "GetCronTab: couldn't find {0}", ctabkey ) );
                return null;
            }
            // Now we've found the right row we expect to find six columns to make up a 
            // crontab entry in D, E, F, G, H, I, J, and then two more columns for start
            // & end in K & L
            string[] flds = new string[6];
            int col = 0;
            for ( ; col < 6; col++)
                flds[col] = GetCellAsString( row, col + 3 );
            string cronex = String.Join( " ", flds );
            double dstart, dend;
            DateTime? start = null; // return nulls if K & L cols are empty
            DateTime? end = null;
            // new DateTime( start.Year, start.Month, start.Day, 23, 59, 59 );
            // If the start and end cells on the cron tab entry on the s2cfg page are time of
            // day eg 09:30:00 and not full date times then they yield DateTime doubles that
            // are < 1.0 as they encode no date/day info. But the Interval arithmetic for the
            // next event in CronManager uses DateTime.Now as a baseline, and that includes
            // date/day info. So we must baseline off the date/day for today too.
            DateTime sod = new DateTime( DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0 );
            double dsod = sod.ToOADate( );
            string sstart = GetCellAsString( row, 3 + col++ );
            string send = GetCellAsString( row, 3 + col++);
            if (Double.TryParse( sstart, out dstart )) {
                if ( dstart < 1.0)
                    dstart += dsod;
                start = DateTime.FromOADate( dstart );
            }
            if (Double.TryParse( send, out dend )) {
                if (dend < 1.0)
                    dend += dsod;
                end = DateTime.FromOADate( dend );
            }
            return new Tuple<String,DateTime?,DateTime?>( cronex, start, end); 
        }

        public String GetWebSock( String wskey ) {
            // We're looking for a row that has 'websock' in the first cell, url in the second,
            // and then wskey in the third.
            int row = FindRow( "websock", "url", wskey );
            if (row == -1) {
                Logr.Log( String.Format( "GetWebSock: couldn't find {0}", wskey ) );
                return null;
            }
            // Now we've found the right row we expect to find three columns to make up a 
            // URL: host, port, path in D, E & F
            string host = GetCellAsString( row, 3);
            string port = GetCellAsString( row, 4);
            string path = GetCellAsString( row, 5);
            if (host == null || port == null || path == null) {
                Logr.Log( String.Format( "GetWebSock: bad host, port or path wskey({0})", wskey));
                return null;
            }
            string url = String.Format( "ws://{0}:{1}/{2}", host, port, path );
            return url;
        }

        public Dictionary<String,String> GetTiingoWebSock( String wskey ) {
            // We're looking for a row that has 'twebsock' in the first cell, tiingo in the second,
            // and then wskey in the third.
            int row = FindRow( "twebsock", "tiingo", wskey );
            if (row == -1) {
                Logr.Log( String.Format( "GetTiingoWebSock: couldn't find {0}", wskey ) );
                return null;
            }
            // Now we've found the right row we expect to find the url in col D
            // and the auth_token on a config row. The auth_token is mandatory
            // for a tiingo websock. 
            string url = GetCellAsString( row, 3 );
            string auth_token = GetQueryConfig( "tiingo", "auth_token" );
            if (url == null || auth_token == null) {
                Logr.Log( String.Format( "GetTiingoWebSock: bad url or auth_token wskey({0})", wskey ) );
                return null;
            }
            var req = new Dictionary<string,string>( ){{"type","twebsock"},{"key",wskey},{"url",url},{"auth_token",auth_token}};
            // Now let's deal with optional elements of a tiingo request: proxy config may be supplied if we have to
            // connect via a proxy.
            GetProxyConfig("tiingo", req);
            return req;
        }

        public void GetProxyConfig(string ctype, Dictionary<string, string> req)
        {
            foreach (string key in s_ProxyKeys)
            {
                string val = GetQueryConfig(ctype, key);
                if (val != "")
                    req.Add(key, val);
            }
        }
    }
}
