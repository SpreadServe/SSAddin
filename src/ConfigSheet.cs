using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

// https://www.quandl.com/api/v1/datasets/WIKI/AAPL.csv
// https://www.quandl.com/api/v1/datasets/WIKI/AAPL.json?trim_start=1985-05-01&trim_end=1997-07-01&sort_order=asc&column=4&collapse=quarterly&transformation=rdiff


namespace SSAddin {

    // Wrapper for all the config info in the s2cfg sheet, if it exists

    class ConfigSheet {

        protected static String s_BaseURL = "https://www.quandl.com/api/v1/datasets";
        protected static Dictionary<string, Func<object, string>> s_QuandlQueryFieldConverters = new Dictionary<string, Func<object, string>>( );

        public ConfigSheet( ) {
            s_QuandlQueryFieldConverters["trim_start"] = ExcelDateNumberToString;
            s_QuandlQueryFieldConverters["trim_end"] = ExcelDateNumberToString;
        }

        public object GetCell( int row, int col ) {
            // TODO: optimise to not create an ExcelReference on every visit.
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

        public String BuildQuandlQuery( Dictionary<string, object> qterms ) {
            if ( !qterms.ContainsKey( "dataset"))
                return "";
            StringBuilder sb = new StringBuilder( s_BaseURL);
            sb.Append( String.Format( "/{0}.csv", qterms["dataset"]));
            qterms.Remove( "dataset");
            string prefix = "?";
            string auth_token = GetQuandlConfig( "auth_token" );
            if (auth_token != "") {
                sb.Append( String.Format( "{0}{1}={2}", prefix, "auth_token", auth_token ) );
                prefix = "&";
            }
            string val;
            foreach ( KeyValuePair<string, object> item in qterms) {
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

        public int FindRow( string c0, string c1, string c2) {
            // We're looking for a row that has c0 in the first cell, c1 in the second,
            // and then c2 in the third.
            int row = 0;
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

        public String GetQuandlQuery( String qkey) {
            // We're looking for a row that has 'quandl' in the first cell, query in the second,
            // and then qkey in the third.
            int row = FindRow( "quandl", "query", qkey);
            if (row == -1) {
                Logr.Log( String.Format( "GetQuandlQuery: couldn't find {0}", qkey));
                return "";
            }
            int col = 3;
            string name;
            object val;
            Dictionary<string, object> qterms = new Dictionary<string, object>( );
            do {
                name = GetCellAsString( row, col );
                val = GetCell( row, col + 1 );
                if (name != null && name != "")
                   qterms.Add( name, val );
                col += 2;
            } while (name != null && name != "");
            return BuildQuandlQuery( qterms );
        }

        public String GetQuandlConfig( String ckey ) {
            // We're looking for a row that has 'quandl' in the first cell, config in the second,
            // and then ckey in the third.
            int row = FindRow( "quandl", "config", ckey );
            if (row == -1) {
                Logr.Log( String.Format( "GetQuandlConfig: couldn't find {0}", ckey));
                return "";
            }
            return GetCellAsString( row, 3 );
        }

        public Tuple<String,DateTime,DateTime> GetCronTab( String ctabkey ) {
            // We're looking for a row that has 'cron' in the first cell, tab in the second,
            // and then ctabkey in the third.
            int row = FindRow( "cron", "tab", ctabkey );
            if (row == -1) {
                Logr.Log( String.Format( "GetQuandlQuery: couldn't find {0}", ctabkey ) );
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
            DateTime start = DateTime.Now;  // default to now
            DateTime end  = new DateTime( start.Year, start.Month, start.Day, 23, 59, 59 );
            // If the start and end cells on the cron tab entry on the s2cfg page are TIME( )
            // fun calls they yield DateTime doubles that are < 1.0 as they encode no date/day.
            // info. But the Interval arithmetic for the next event in CronManager uses
            // DateTime.Now as a baseline, and that includes date/day info. So we must
            // baseline off the date/day for today too.
            DateTime sod = new DateTime( start.Year, start.Month, start.Day, 0, 0, 0 );
            double dsod = sod.ToOADate( );
            string sstart = GetCellAsString( row, 3 + col++ );
            string send = GetCellAsString( row, 3+col++);
            // Add code here to use ToOADate to baseline against today
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
            return new Tuple<String,DateTime,DateTime>( cronex, start, end); 
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
    }
}
