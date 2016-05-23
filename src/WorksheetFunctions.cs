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
        static double       s_SSAddinVersion = 0.1;
		#endregion

        #region Regular worksheet functions
        /*
        [ExcelFunction( Description = "Runs a blockspring function", Category = "Blockspring" )]
        public static object BLOCKSPRING( [ExcelArgument( Name = "Block ID", Description = "The block to run" )] string block_id, [ExcelArgument( Name = "Key", Description = "The key of the input you're providing" )] string key, [ExcelArgument( Name = "Range", Description = "The range of data you want to pull" )] object range, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey1, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange1, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey2, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange2, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey3, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange3, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey4, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange4, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey5, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange5, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey6, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange6, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey7, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange7, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey8, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange8, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey9, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange9, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey10, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange10, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey11, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange11, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey12, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange12, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey13, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange13, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey14, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange14, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey15, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange15, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey16, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange16, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey17, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange17, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey18, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange18, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey19, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange19, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey20, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange20, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey21, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange21, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey22, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange22, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey23, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange23, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey24, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange24, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey25, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange25, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey26, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange26, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey27, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange27, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey28, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange28, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey29, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange29, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey30, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange30, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey31, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange31, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey32, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange32, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey33, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange33, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey34, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange34, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey35, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange35, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey36, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange36, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey37, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange37, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey38, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange38, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey39, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange39, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey40, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange40, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey41, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange41, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey42, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange42, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey43, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange43, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey44, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange44, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey45, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange45, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey46, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange46, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey47, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange47, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey48, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange48, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey49, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange49, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey50, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange50, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey51, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange51, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey52, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange52, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey53, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange53, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey54, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange54, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey55, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange55, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey56, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange56, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey57, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange57, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey58, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange58, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey59, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange59, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey60, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange60, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey61, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange61, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey62, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange62, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey63, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange63, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey64, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange64, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey65, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange65, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey66, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange66, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey67, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange67, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey68, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange68, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey69, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange69, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey70, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange70, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey71, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange71, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey72, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange72, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey73, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange73, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey74, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange74, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey75, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange75, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey76, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange76, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey77, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange77, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey78, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange78, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey79, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange79, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey80, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange80, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey81, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange81, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey82, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange82, [ExcelArgument( Name = "Optional Key", Description = "An optional key of the input you're providing" )] string optionalKey83, [ExcelArgument( Name = "Optional Range", Description = "An optional range of data you want to pull" )] object optionalRange83 ) {
            return ExcelError.ExcelErrorNA;
        } */

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
            if (s_WebClient.AddRequest( "quandl", qkey, url ))
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
            if (s_WebClient.AddRequest( "tiingo", qkey, url, auth_token ))
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
            if (s_WebClient.AddRequest( "websock", wskey, url))
                return s_Submitted;
            return ExcelError.ExcelErrorValue;
        }

        [ExcelFunction( Description = "Connect to tiingo web socket." )]
        public static object s2twebsock( 
            [ExcelArgument( Name = "SockKey", Description = "twebsock url key in s2cfg!C" )] string wskey ) {
            Tuple<String,String> urlauth = s_ConfigSheet.GetTiingoWebSock( wskey);
            if (urlauth == null) {
                return ExcelMissing.Value;
            }
            if (s_WebClient.AddRequest( "twebsock", wskey, urlauth.Item1, urlauth.Item2 ))
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

        [ExcelFunction( Description = "Volatile: pull data from S2 web socket cache." )]
        public static object s2vwscache(
            [ExcelArgument( Name = "QueryKey", Description = "websock query key in s2cfg!C" )] string wkey,
            [ExcelArgument( Name = "CellKey", Description = "m2_6_0 for col 3, row 7 on first sheet" )] string ckey) {
            return s2wscache( wkey, ckey, null );
        }

        #endregion

        #region RTD functions
        [ExcelFunction( Description = "RTD: Subscribe to properties of S2 cache." )]
        public static object s2sub(
            [ExcelArgument( Name = "SubCache", Description = "[quandl|tiingo|cron|websock]" )] string subcache,
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
