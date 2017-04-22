// Copyright © Babbington Slade

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using ExcelDna.Integration;

// Some notes on threading: there are multiple threads in play in this code, so locking is used fairly
// freely in the code, and #region/#endregion markers are used to indicate which threads they're running on.
// There are three kinds of threads...
// 1. The Excel thread: the worksheet functions in this .cs run on the Excel thread.
// 2. The background worker thread: this thread is launched in SSWebClient. It's top 
//    level loop is in SSWebClient.BackgroundWork( ). SSWebClient.AddRequest( ) is used
//    by the Excel thread to pass work via a queue to this thread. This thread handles
//    HTTP GET style queries synchronously, so we avoid blocking the Excel thread. It 
//    also initiates the websock and tiingo websock objects.
// 3. .Net pool threads: .Net dispatches all kinds of events on pool threads: timers and
//    socket events. Consequently a lot of the cron code as well as websock and tiingo
//    websock callbacks run on pool threads. They dump results in caches that the Excel
//    thread accesses, so locking is needed. They also send results back to Excel via
//    the RTDServer; Excel is the RTD client. The RTDServer has methods that are invoked on
//    the Excel thread, so locking is necessary there too.
// JOS 2016-05-25

namespace SSAddin {
	public static class WorksheetFunctions {
		#region Fields
        static ConfigSheet  s_ConfigSheet = new ConfigSheet( );
        static DataCache    s_Cache = DataCache.Instance( );
        static SSWebClient  s_WebClient = SSWebClient.Instance( );
        static CronManager  s_CronMgr = CronManager.Instance( );
        static string       s_Submitted = "OK";
        static double       s_SSAddinVersion = 0.1;
		#endregion

        #region Regular worksheet functions

        [ExcelFunction( Description = "Version info." )]
        public static object s2about( ) {
            return String.Format( "SSAddin {0} Excel {1}", s_SSAddinVersion, ExcelDnaUtil.ExcelVersion );
        }

        [ExcelFunction( Description = "Launch quandl query.")]
        public static object s2quandl( 
            [ExcelArgument( Name = "QueryKey", Description = "quandl query key in s2cfg!C" )] string qkey,
            [ExcelArgument( Name = "Trigger", Description = "dummy to trigger recalc" )] object trigger )
		{
            String url = s_ConfigSheet.GetQueryURL( "quandl", qkey );
            if ( url == null || url == "") {
                return ExcelMissing.Value;
            }
            var req = new Dictionary<string,string>(){ {"type","quandl"}, {"key",qkey},{"url",url}};
            s_ConfigSheet.GetProxyConfig("quandl", req);
            if (s_WebClient.AddRequest( req))
                return s_Submitted;
            return ExcelError.ExcelErrorGettingData;
		}

        [ExcelFunction( Description = "Launch tiingo query." )]
        public static object s2tiingo(
            [ExcelArgument( Name = "QueryKey", Description = "tiingo query key in s2cfg!C" )] string qkey,
            [ExcelArgument( Name = "Trigger", Description = "dummy to trigger recalc" )] object trigger ) {
            String url = s_ConfigSheet.GetQueryURL( "tiingo", qkey );
            if (url == null || url == "") {
                return ExcelMissing.Value;
            }
            // quandl puts the auth_token in the URL. tiingo is different, and puts it in the HTTP headers
            String auth_token = s_ConfigSheet.GetQueryConfig( "tiingo", "auth_token" );
            if (auth_token == null || auth_token == "") {
                return ExcelMissing.Value;
            }
            var req = new Dictionary<string,string>(){ {"type","tiingo"}, {"key",qkey},{"url",url},{"auth_token",auth_token}};
            s_ConfigSheet.GetProxyConfig("tiingo", req);
            if (s_WebClient.AddRequest( req))
                return s_Submitted;
            return ExcelError.ExcelErrorGettingData;
        }

        [ExcelFunction(Description = "Launch baremetrics query.")]
        public static object s2baremetrics(
            [ExcelArgument(Name = "QueryKey", Description = "baremetrics query key in s2cfg!C")] string qkey,
            [ExcelArgument(Name = "Trigger", Description = "dummy to trigger recalc")] object trigger)
        {
            String url = s_ConfigSheet.GetQueryURL("baremetrics", qkey);
            if (url == null || url == "")
            {
                return ExcelMissing.Value;
            }
            // baremetrics puts the auth_token in the HTTP headers
            String auth_token = s_ConfigSheet.GetQueryConfig("baremetrics", "auth_token");
            if (auth_token == null || auth_token == "")
            {
                return ExcelMissing.Value;
            }
            var req = new Dictionary<string, string>() { { "type", "baremetrics" }, { "key", qkey }, { "url", url }, { "auth_token", auth_token } };
            s_ConfigSheet.GetProxyConfig("baremetrics", req);
            if (s_WebClient.AddRequest(req))
                return s_Submitted;
            return ExcelError.ExcelErrorGettingData;
        }

        [ExcelFunction( Description = "Schedule cron job." )]
        public static object s2cron(
            [ExcelArgument( Name = "CronKey", Description = "cron tab key in s2cfg!C" )] string ckey)
        {
            Tuple<String,DateTime?,DateTime?> tup = s_ConfigSheet.GetCronTab( ckey );
            if (tup == null) {
                return ExcelMissing.Value;
            }
            if (s_CronMgr.AddCron( ckey, tup.Item1, tup.Item2, tup.Item3))
                return s_Submitted;
            return ExcelError.ExcelErrorValue;
        }

        [ExcelFunction( Description = "Connect to web socket." )]
        public static object s2websock(
            [ExcelArgument( Name = "SockKey", Description = "websock url key in s2cfg!C" )] string wskey ) {
            String url = s_ConfigSheet.GetWebSock( wskey );
            if (url == null) {
                return ExcelMissing.Value;
            }
            if (s_WebClient.AddRequest( new Dictionary<string,string>(){ {"type","websock"}, {"key",wskey},{"url",url}}))
                return s_Submitted;
            return ExcelError.ExcelErrorValue;
        }

        [ExcelFunction( Description = "Connect to tiingo web socket." )]
        public static object s2twebsock( 
            [ExcelArgument( Name = "SockKey", Description = "twebsock url key in s2cfg!C" )] string wskey ) {
            Dictionary<string,string> req = s_ConfigSheet.GetTiingoWebSock( wskey);
            if ( req == null) {
                return ExcelMissing.Value;
            }
            if (s_WebClient.AddRequest( req))
                return s_Submitted;
            return ExcelError.ExcelErrorValue;
        }

        [ExcelFunction( Description = "Pull data from S2 quandl cache.")]
		public static object s2qcache( 
            [ExcelArgument(Name="QueryKey", Description="quandl query key in s2cfg!C")] string qkey,
            [ExcelArgument(Name="XOffset", Description="column offset to cache position. 0 default")] int xoffset,
            [ExcelArgument(Name="YOffset", Description="row offset to cache position. 0 default" )] int yoffset,
            [ExcelArgument(Name="Trigger", Description="dummy to trigger recalc")] object trigger)
		{
            if ( !s_Cache.ContainsQuandlKey( qkey)) {
                return ExcelMissing.Value;
            }
            // Figure out our caller's posn in the sheet; that's the cell we'll pull from the cache.
            // If offsets are supplied use them to calc cell posn too. xoffset & yoffset will default
            // to 0 if not supplied in the sheet. 
            ExcelReference caller = XlCall.Excel( XlCall.xlfCaller) as ExcelReference;
            string val = s_Cache.GetQuandlCell( qkey, caller.RowFirst-yoffset, caller.ColumnFirst-xoffset);
            if ( val == null ) {
                return ExcelError.ExcelErrorNA; //  ExcelMissing.Value;
            }
            return val;
		}

        [ExcelFunction( Description = "Volatile: pull data from S2 quandl cache.", IsVolatile = true )]
        public static object s2vqcache(
            [ExcelArgument( Name = "QueryKey", Description = "quandl query key in s2cfg!C" )] string qkey,
            [ExcelArgument( Name = "XOffset", Description = "column offset to cache position. 0 default" )] int xoffset,
            [ExcelArgument( Name = "YOffset", Description = "row offset to cache position. 0 default" )] int yoffset) {
            return s2qcache( qkey, xoffset, yoffset, null );
        }

        [ExcelFunction( Description = "Pull data from S2 tiingo cache." )]
        public static object s2tcache(
            [ExcelArgument( Name = "QueryKey", Description = "tiingo query key in s2cfg!C" )] string qkey,
            [ExcelArgument( Name = "XOffset", Description = "column offset to cache position. 0 default" )] int xoffset,
            [ExcelArgument( Name = "YOffset", Description = "row offset to cache position. 0 default" )] int yoffset,
            [ExcelArgument( Name = "Trigger", Description = "dummy to trigger recalc" )] object trigger ) {
            if (!s_Cache.ContainsTiingoKey( qkey )) {
                return ExcelMissing.Value;
            }
            // Figure out our caller's posn in the sheet; that's the cell we'll pull from the cache.
            // If offsets are supplied use them to calc cell posn too. xoffset & yoffset will default
            // to 0 if not supplied in the sheet. 
            ExcelReference caller = XlCall.Excel( XlCall.xlfCaller ) as ExcelReference;
            string val = s_Cache.GetTiingoCell( qkey, caller.RowFirst - yoffset, caller.ColumnFirst - xoffset );
            if (val == null) {
                return ExcelError.ExcelErrorNA; //  ExcelMissing.Value;
            }
            return val;
        }

        [ExcelFunction( Description = "Volatile: pull data from S2 tiingo cache.", IsVolatile = true )]
        public static object s2vtcache(
            [ExcelArgument( Name = "QueryKey", Description = "quandl query key in s2cfg!C" )] string qkey,
            [ExcelArgument( Name = "XOffset", Description = "column offset to cache position. 0 default" )] int xoffset,
            [ExcelArgument( Name = "YOffset", Description = "row offset to cache position. 0 default" )] int yoffset ) {
            return s2tcache( qkey, xoffset, yoffset, null );
        }

        [ExcelFunction( Description = "Pull data from S2 baremetrics cache." )]
        public static object s2bcache(
            [ExcelArgument( Name = "QueryKey", Description = "baremetrics query key in s2cfg!C" )] string qkey,
            [ExcelArgument( Name = "Date", Description = "date key into result set. Use s2today, not Excel's TODAY" )] int xldate,
            [ExcelArgument( Name = "Field", Description = "field in the result set" )] string field,
            [ExcelArgument( Name = "Trigger", Description = "dummy to trigger recalc" )] object trigger ) {
            if (!s_Cache.ContainsBareKey( qkey )) {
                return ExcelMissing.Value;
            }
            string val = s_Cache.GetBareField( qkey, s_ConfigSheet.ExcelDateNumberToString( xldate), field);
            if (val == null) {
                return ExcelError.ExcelErrorNA; //  ExcelMissing.Value;
            }
            return val;
        }

        [ExcelFunction( Description = "Volatile: pull data from S2 baremetrics cache.", IsVolatile = true )]
        public static object s2vbcache(
            [ExcelArgument( Name = "QueryKey", Description = "baremetrics query key in s2cfg!C" )] string qkey,
            [ExcelArgument( Name = "Date", Description = "date key into result set. Use s2today, not Excel's TODAY" )] int xldate,
            [ExcelArgument( Name = "Field", Description = "field in the result set" )] string field ) {
            return s2bcache( qkey, xldate, field, null );
        }

        [ExcelFunction( Description = "Non volatile alternate to Excel's TODAY.")]
        public static object s2today(
            [ExcelArgument( Name = "Offset", Description = "days +/- from today" )] int offset,
            [ExcelArgument( Name = "Trigger", Description = "dummy to trigger recalc" )] object trigger ) {
            DateTime dt = DateTime.Now.AddDays( Convert.ToDouble( offset ) );
            return dt.ToString( "yyyy-MM-dd" );
        }

        [ExcelFunction( Description = "Pull data from S2 web socket cache." )]
        public static object s2wscache(
            [ExcelArgument( Name = "QueryKey", Description = "websock query key in s2cfg!C" )] string wkey,
            [ExcelArgument( Name = "CellKey", Description = "m2_6_0 for col 3, row 7 on first sheet" )] string ckey,
            [ExcelArgument( Name = "Trigger", Description = "dummy to trigger recalc" )] object trigger ) {
            // Figure out our caller's posn in the sheet; that's the cell we'll pull from the cache.
            // If offsets are supplied use them to calc cell posn too. xoffset & yoffset will default
            // to 0 if not supplied in the sheet. 
            string val = s_Cache.GetWSCell( wkey, ckey);
            if (val == null) {
                return ExcelError.ExcelErrorNA; //  ExcelMissing.Value;
            }
            return val;
        }

        [ExcelFunction( Description = "Volatile: pull data from S2 web socket cache.", IsVolatile = true)]
        public static object s2vwscache(
            [ExcelArgument( Name = "QueryKey", Description = "websock query key in s2cfg!C" )] string wkey,
            [ExcelArgument( Name = "CellKey", Description = "m2_6_0 for col 3, row 7 on first sheet" )] string ckey) {
            return s2wscache( wkey, ckey, null );
        }

        #endregion

        #region RTD functions
        [ExcelFunction( Description = "RTD: Subscribe to properties of S2 cache." )]
        public static object s2sub(
            [ExcelArgument( Name = "SubCache", Description = "[quandl|tiingo|cron|websock|twebsock]" )] string subcache,
            [ExcelArgument(Name="CacheKey", Description="Row key from s2cfg")] string ckey,
            [ExcelArgument(Name="Property", Description="[status|count|next|last|mX_Y_Z|ticker_field]")] string prop)
        {
            string[] arrey = { subcache, ckey, prop};
            string stopic = String.Join( ".", arrey);
            Logr.Log( String.Format( "s2sub: {0}", stopic));
            // Send a message to the worker thread about this subscription. It may need to fwd
            // to another object eg TWSCallback for subscription management.
            var sdict = new Dictionary<string, string>() { { "type", "s2sub" }, { "key", stopic },
                                    { "subcache", subcache }, { "ticker_field", prop}, { "cachekey", ckey} };
            s_WebClient.AddRequest( sdict);
            // Make the RTD call to Excel or SpreadServe's internal RTD API to let it know
            // about the new subscription.
            return XlCall.RTD( "SSAddin.RTDServer", null, stopic);
        }
        #endregion
	}
}
