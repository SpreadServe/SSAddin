// Copyright © Babbington Slade

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using ExcelDna.Integration;

namespace SSAddin {
	public static class WorksheetFunctions {
		#region Fields
        static ConfigSheet  s_ConfigSheet = new ConfigSheet( );
        static DataCache    s_Cache = DataCache.Instance( );
        static SSWebClient  s_WebClient = SSWebClient.Instance( );
        static CronManager  s_CronMgr = CronManager.Instance( );
        static string       s_Submitted = "OK";
		#endregion

        #region Regular worksheet functions

        [ExcelFunction( Description = "Version info." )]
        public static object s2about( ) {
            return "SSAddin beta 0.1";
        }

        [ExcelFunction( Description = "Launch quandl query.")]
        public static object s2quandl( 
            [ExcelArgument( Name = "QueryKey", Description = "quandl query key in s2cfg!C" )] string qkey,
            [ExcelArgument( Name = "Trigger", Description = "dummy to trigger recalc" )] object trigger )
		{
            String url = s_ConfigSheet.GetQuandlQuery( qkey );
            if ( url == null || url == "") {
                return ExcelMissing.Value;
            }
            if (s_WebClient.AddRequest( "quandl", qkey, url ))
                return s_Submitted;
            return ExcelError.ExcelErrorGettingData;
		}

        [ExcelFunction( Description = "Schedule cron job." )]
        public static object s2cron(
            [ExcelArgument( Name = "CronKey", Description = "cron tab key in s2cfg!C" )] string ckey)
        {
            Tuple<String,DateTime,DateTime> tup = s_ConfigSheet.GetCronTab( ckey );
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
            if (s_WebClient.AddRequest( "websock", wskey, url))
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
        #endregion

        #region RTD functions
        [ExcelFunction( Description = "Subscribe to properties of S2 cache." )]
        public static object s2sub(
            [ExcelArgument( Name = "SubCache", Description = "[quandl|cron|websock]" )] string subcache,
            [ExcelArgument(Name="CacheKey", Description="Row key from s2cfg")] string ckey,
            [ExcelArgument(Name="Property", Description="[status|count|next|last|mX_Y_Z]")] string prop)
        {
            string[] arrey = { subcache, ckey, prop};
            string stopic = String.Join( ".", arrey);
            Logr.Log( String.Format( "s2sub: {0}", stopic));
            return XlCall.RTD( "SSAddin.RTDServer", null, stopic);
        }
        #endregion
	}
}
