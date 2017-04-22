using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Net;
using System.Diagnostics;
using NCrontab;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SSAddin {
    #region Deserialization classes
    // tiingo historical data format
    class SSTiingoHistPrice {
        public string   date { get; set; }
        public float    open { get; set; }
        public float    close { get; set; }
        public float    high { get; set; }
        public float    low { get; set; }
        public float    volume { get; set; }
        public float    adjOpen { get; set; }
        public float    adjClose { get; set; }
        public float    adjHigh { get; set; }
        public float    adjLow { get; set; }
        public float    adjVolume { get; set; }
    }
    // baremetrics metrics query responses
    // https://developers.baremetrics.com/reference#available-metrics
    class BareMetricsSummaryArrayElement {
        public string   human_date { get; set; }
        public string   date { get; set; }
        public int      active_customers { get; set; }
        public int      active_subscriptions { get; set; }
        public int      add_on_mrr { get; set; }
        public int      arpu { get; set; }
        public int      arr { get; set; }
        public int      cancellations { get; set; }
        public int      coupons { get; set; }
        public int      downgrades { get; set; }
        public int      failed_charges { get; set; }
        public int      fees { get; set; }
        public int      ltv { get; set; }
        public int      mrr { get; set; }
        public int      net_revenue { get; set; }
        public int      new_customers { get; set; }
        public int      other_revenue { get; set; }
        public int      reactivated_customers { get; set; }
        public int      refunds { get; set; }
        public int      revenue_churn { get; set; }
        public int      trial_conversions { get; set; }
        public int      upgrades { get; set; }
        public int      user_churn { get; set; }
    }

    class BareMetricsSummary {
        public List<BareMetricsSummaryArrayElement> metrics { get; set; }
    }

    #endregion

    class SSWebClient {
        protected static DataCache      s_Cache = DataCache.Instance( );
        protected static char[]         csvDelimiterChars = { ',' };
        protected static SSWebClient    s_Instance;
        protected static object         s_InstanceLock = new object( );

        protected Queue<Dictionary<string,string>>  m_InputQueue;
        protected HashSet<String>                   m_InFlight;
        protected Dictionary<String, WSCallback>    m_WSCallbacks;
        protected TWSCallback                       m_TWSCallback;
        protected List<Dictionary<string, string>>  m_PendingTiingoSubs;

        protected Thread            m_WorkerThread;     // for executing the web query
        protected ManualResetEvent  m_Event;            // control worker thread sleep
        protected String            m_TempDir;
        protected int               m_QuandlCount;      // total number of quandl queries
        protected int               m_TiingoCount;      // total number of tiingo queries
        protected int               m_BareCount;        // total number of baremetrics queries


        #region Excel thread methods

        protected SSWebClient( )
        {
            m_TempDir = System.IO.Path.GetTempPath( );
            m_InputQueue = new Queue<Dictionary<string,string>>( );
            m_InFlight = new HashSet<String>( );
            m_WSCallbacks = new Dictionary<string, WSCallback>( );
            m_Event = new ManualResetEvent( false);
            m_PendingTiingoSubs = new List<Dictionary<string, string>>( );
            m_WorkerThread = new Thread( BackgroundWork );
            m_WorkerThread.Start( );
            m_QuandlCount = 0;
            m_TiingoCount = 0;
            m_BareCount = 0;

            // Push out an RTD update for the overall quandl query count. This will mean
            // that trigger parms driven by quandl.all.count don't have #N/A as input for
            // long, which will enable eg ycb_pub_quandl.xls qlPiecewiseYieldCurve to
            // calc almost immediately. JOS 2015-07-29
            UpdateRTD( "quandl", "all", "count", String.Format( "{0}", m_QuandlCount++));
            UpdateRTD( "tiingo", "all", "count", String.Format( "{0}", m_TiingoCount++ ) );
            UpdateRTD( "baremetrics", "all", "count", String.Format("{0}", m_BareCount++));
        }

        public static SSWebClient Instance( ) {
            // Unlikley that two threads will attempt to instance this singleton at the
            // same time, but we'll lock just in case.
            lock (s_InstanceLock) {
                if (s_Instance == null) {
                    s_Instance = new SSWebClient( );
                }
                return s_Instance;
            }
        }


        public bool AddRequest( Dictionary<string, string> req) {
            // Every request *must* have type, key, url. There may be other optionals like
            // auth_token, https_proxy_host...
            string type = req["type"];
            string key = req["key"];
            string fkey = String.Format( "{0}.{1}", req["type"], req["key"]);
            bool isWebQuery = (type == "quandl" || type == "tiingo" || type == "baremetrics");
            // Is this job pending or in progress?
            lock (m_InFlight) {
                if (m_InFlight.Contains( fkey )) {   // Queued or running...
                    if ( isWebQuery)
                        Logr.Log( String.Format( "~A AddRequest: {0} is already inflight", fkey ) );
                    return false;                   // so bail
                }
                // Nested locking - look out! We're on the Excel thread here as we're invoked by
                // worksheet functions. The background worker thread does use this lock too, but
                // not at the same time as m_InFlight, so we should be OK.
                lock (m_InputQueue) {
                    // Running on the main Excel thread here. Q the work, and
                    // signal the background thread to wake up...
                    if ( isWebQuery)
                        Logr.Log( String.Format( "~A AddRequest adding {0} {1}", type, key ) );
                    m_InputQueue.Enqueue( req);
                }
                // NB some fkeys are only ever added to m_InFlight, and are never removed. For
                // instance the s2sub notifications. We only need an s2sub notification to go to
                // the background thread once to create websock subscriptions. And if it's an
                // hbcount for instance, there's no work on the background thread, so we'll
                // let it through once, then block subsequent notifies. We may want to revisit
                // this logic in future if we need to resubscribe to websock topics. But for
                // the time being let's work around by starting and stopping the sheet.
                // JOS 2016-05-26
                m_InFlight.Add( fkey);
                m_Event.Set( ); // signal worker thread to wake
            }
            return true;
        }

        #endregion

        #region Pool thread methods

        protected void WSCallbackClosed( string wskey) {
            lock (m_InFlight) {
                // removing the key of the websock allows another incoming request
                // with the same key
                m_InFlight.Remove( wskey );
                m_WSCallbacks.Remove( wskey );
            }
        }

        protected void TWSCallbackClosed( string wskey ) {
            lock (m_InFlight) {
                // releasing the existing tiingo websock callback handler enables the 
                // BackgroundWork method to create another one
                m_InFlight.Remove( wskey );
                m_TWSCallback = null;
            }
        }

        #endregion

        #region Worker thread methods

        protected Dictionary<string,string> GetWork( ) {
            // Put this oneliner in its own method to wrap the locking. We can't
            // hold the lock while we're looping in BackgroundWork( ) as that
            // will prevent the Excel thread adding new requests.
            lock ( m_InputQueue) {
                if (m_InputQueue.Count > 0)
                    return m_InputQueue.Dequeue( );
                return null;
            }
        }

        public void BackgroundWork( ) {
            // We're running on the background thread. Loop until we're told to exit...
            Logr.Log( String.Format( "~A BackgroundWork thread started"));
            // Main loop for worker thread. It will briefly hold the m_InFlight lock when
            // it removes entries, and also the m_InputQueue lock in GetWork( ) as it
            // removes entries. DoQuandlQuery( ) will grab the cache lock when it's
            // adding cache entries. Obviously, no lock should be held permanently! JOS 2015-04-31
            while (true) {
                // Wait for a signal from the other thread to say there's some work.
                m_Event.WaitOne( );
                Dictionary<string,string> work = GetWork( );
                while ( work != null) {
                    if (work["type"] == "stop") {
                        // exit req from excel thread
                        Logr.Log( String.Format( "~A BackgroundWork thread exiting" ) );
                        return;
                    }
                    string fkey = String.Format("{0}.{1}", work["type"], work["key"]);
                    Logr.Log(String.Format("~A BackgroundWork new request fkey({0})", fkey));

                    if (work["type"] == "quandl") {
                        // run query synchronously here on background worker thread
                        bool ok = DoQuandlQuery( work);
                        // query done, so remove key from inflight, which will permit
                        // the query to be resubmitted
                        lock (m_InFlight) {
                            m_InFlight.Remove( fkey );
                        }
                    }
                    else if (work["type"] == "tiingo") {
                        // run query synchronously here on background worker thread
                        bool ok = DoTiingoQuery( work );
                        // query done, so remove key from inflight, which will permit
                        // the query to be resubmitted
                        lock (m_InFlight) {
                            m_InFlight.Remove( fkey );
                        }
                    }
                    else if (work["type"] == "baremetrics")
                    {
                        // run query synchronously here on background worker thread
                        bool ok = DoBareQuery(work);
                        // query done, so remove key from inflight, which will permit
                        // the query to be resubmitted
                        lock (m_InFlight)
                        {
                            m_InFlight.Remove(fkey);
                        }
                    }
                    else if (work["type"] == "websock") {
                        WSCallback wscb = new WSCallback( work, this.WSCallbackClosed );
                        // We don't want to remove the inflight key here as there will be
                        // async callbacks to WSCallback on pool threads when updates
                        // arrive on the SS websock. So we leave the key in place to
                        // prevent AddRequest, which is on the Excel thread, creating 
                        // a request for a new WSCallback. 
                        lock (m_InFlight) {
                            m_WSCallbacks.Add( fkey, wscb );
                        }
                    }
                    else if (work["type"] == "twebsock") {
                        // We don't want to remove the inflight key here as there will be
                        // async callbacks to TWSCallback on pool threads when updates
                        // arrive on the tiingo websock. So we leave the key in place to
                        // prevent AddRequest, which is on the Excel thread, creating 
                        // a request for a new TWSCallback. 
                        lock (m_InFlight) {
                            if (m_TWSCallback == null) {
                                m_TWSCallback = new TWSCallback( work, this.TWSCallbackClosed );
                                if (m_PendingTiingoSubs.Count > 0) {
                                    m_TWSCallback.AddSubscriptions(m_PendingTiingoSubs);
                                    m_PendingTiingoSubs.Clear();
                                }
                            }
                        }
                    }
                    else if (work["type"] == "s2sub") {
                        if (work["subcache"] == "twebsock") {
                            // New subscription to a tiingo websock. If the TWSCallback
                            // doesn't exist yet cache it, but if it does pass it through
                            lock (m_InFlight) {
                                m_PendingTiingoSubs.Add(work);
                                if (m_TWSCallback != null) {
                                    m_TWSCallback.AddSubscriptions(m_PendingTiingoSubs);
                                    m_PendingTiingoSubs.Clear();
                                }
                            }
                        }
                    }
                    work = GetWork( );
                }
                // We've exhausted the queued work, so reset the event so that we wait in the
                // WaitOne( ) invocation above until another thread signals that there's some
                // more work.
                m_Event.Reset( );
            }
        }

        protected void UpdateRTD( string subcache, string qkey, string subelem, string value ) {
            // The RTD server doesn't necessarily exist. If no cell calls 
            // s2sub( ) it won't be instanced by Excel.
            RTDServer rtd = RTDServer.GetInstance( );
            if ( rtd == null)
                return;
            string stopic = String.Format( "{0}.{1}.{2}", subcache, qkey, subelem );
            rtd.CacheUpdate( stopic, value );
        }

        protected void ConfigureProxy(Dictionary<string, string> work, WebClient wc)
        {
            // If the dictionary has proxy config, then set it up...
            if (!work.ContainsKey("http_proxy_host"))
                return;

            int port = 80;
            string host = work["http_proxy_host"];
            if (work.ContainsKey("http_proxy_port"))
            {
                if (!Int32.TryParse(work["http_proxy_port"], out port))
                    port = 80;
            }

            WebProxy proxy = new WebProxy( String.Format("{0}:{1}", host, port), true);
            string user = "", pass = "";
            if (work.ContainsKey("http_proxy_user") && work.ContainsKey("http_proxy_password")) {
                user = work["http_proxy_user"];
                pass = work["http_proxy_password"];
                proxy.Credentials = new NetworkCredential( user, pass);
            }
            wc.Proxy = proxy;
            Logr.Log(String.Format("ConfigureProxy host({0}) port({1}) user({2}) pass({3})", host, port, user, pass));
        }

        protected bool DoQuandlQuery( Dictionary<string,string> work)
		{
            string qkey = work["key"];
            string url = work["url"];
            string line = "";
            string lineCount = "0";
			try	{
                // Set up the web client to HTTP GET
                var client = new WebClient( );
                ConfigureProxy(work, client);
                Stream data = client.OpenRead( url);
                var reader = new StreamReader( data);
                // Local file to dump result
                int pid = Process.GetCurrentProcess( ).Id;
                string csvfname = String.Format( "{0}\\{1}_{2}.csv", m_TempDir, qkey, pid );
                Logr.Log( String.Format( "running quandl qkey({0}) {1} persisted at {2}", qkey, url, csvfname));
                var csvf = new StreamWriter( csvfname );
                UpdateRTD( "quandl", qkey, "status", "starting" );
                // Clear any previous result from the cache so we don't append repeated data
                s_Cache.ClearQuandl( qkey );
                while ( reader.Peek( ) >= 0) {
                    // For each CSV line returned by quandl, dump to localFS, add to in mem cache, and 
                    // send a line count update to any RTD subscriber
                    line = reader.ReadLine( );
                    csvf.WriteLine( line );
                    lineCount = String.Format( "{0}", s_Cache.AddQuandlLine( qkey, line.Split( csvDelimiterChars)));
                    UpdateRTD( "quandl", qkey, "count", lineCount );
                }
                csvf.Close( );
                data.Close();
                reader.Close();
                UpdateRTD( "quandl", qkey, "status", "complete" );
                UpdateRTD( "quandl", "all", "count", String.Format( "{0}", m_QuandlCount++ ) );
                Logr.Log( String.Format( "quandl qkey({0}) complete count({1})", qkey, lineCount));
                return true;
			}
			catch( System.IO.IOException ex) {
                Logr.Log( String.Format( "quandl qkey({0}) url({1}) {2}", qkey, url, ex) );
			}
            catch (System.Net.WebException ex) {
                Logr.Log( String.Format( "quandl qkey({0}) url({1}) {2}", qkey, url, ex ) );
            }
            return false;
		}

        protected bool DoTiingoQuery( Dictionary<string,string> work) {
            string qkey = work["key"];
            string url = work["url"];
            string auth_token = work["auth_token"];
            string line = "";
            string lineCount = "0";
            try {
                // Set up the web client to HTTP GET
                var client = new WebClient( );
                ConfigureProxy(work, client);
                client.Headers.Set( "Content-Type", "application/json" );
                client.Headers.Set( "Authorization", String.Format("Token {0}", auth_token ));
                Stream data = client.OpenRead( url );
                var reader = new StreamReader( data );
                // Local file to dump result
                int pid = Process.GetCurrentProcess( ).Id;
                string jsnfname = String.Format( "{0}\\{1}_{2}.jsn", m_TempDir, qkey, pid );
                Logr.Log( String.Format( "running tiingo qkey({0}) {1} persisted at {2}", qkey, url, jsnfname ) );
                var jsnf = new StreamWriter( jsnfname );
                UpdateRTD( "tiingo", qkey, "status", "starting" );
                // Clear any previous result from the cache so we don't append repeated data
                s_Cache.ClearTiingo( qkey );
                StringBuilder sb = new StringBuilder( );
                while (reader.Peek( ) >= 0) {
                    // For each json line returned by tiingo, dump to localFS, add to in mem cache
                    line = reader.ReadLine( );
                    jsnf.WriteLine( line );
                    sb.AppendLine( line );
                }
                jsnf.Close( );
                data.Close( );
                reader.Close( );
                UpdateRTD( "tiingo", qkey, "status", "complete" );
                UpdateRTD( "tiingo", "all", "count", String.Format( "{0}", m_TiingoCount++ ) );
                Logr.Log( String.Format( "tiingo qkey({0}) complete count({1})", qkey, lineCount ) );
                List<SSTiingoHistPrice> updates = JsonConvert.DeserializeObject<List<SSTiingoHistPrice>>( sb.ToString( ));
                s_Cache.UpdateTHPCache( qkey, updates );
                UpdateRTD( "tiingo", qkey, "count", String.Format( "{0}", updates.Count) );
                return true;
            }
            catch (System.IO.IOException ex) {
                Logr.Log( String.Format( "tiingo qkey({0}) url({1}) {2}", qkey, url, ex ) );
            }
            catch (System.Net.WebException ex) {
                Logr.Log( String.Format( "tiingo  qkey({0}) url({1}) {2}", qkey, url, ex ) );
            }
            return false;
        }

        protected bool DoBareQuery(Dictionary<string, string> work)
        {
            string qkey = work["key"];
            string url = work["url"];
            string auth_token = work["auth_token"];
            string line = "";
            string lineCount = "0";
            try
            {
                // Set up the web client to HTTP GET
                var client = new WebClient();
                ConfigureProxy(work, client);
                client.Headers.Set("Content-Type", "application/json");
                client.Headers.Set("Authorization", String.Format("Bearer {0}", auth_token));
                Stream data = client.OpenRead(url);
                var reader = new StreamReader(data);
                // Local file to dump result
                int pid = Process.GetCurrentProcess().Id;
                string jsnfname = String.Format("{0}\\{1}_{2}.jsn", m_TempDir, qkey, pid);
                Logr.Log(String.Format("running baremetric qkey({0}) {1} persisted at {2}", qkey, url, jsnfname));
                var jsnf = new StreamWriter(jsnfname);
                UpdateRTD("baremetrics", qkey, "status", "starting");
                // Clear any previous result from the cache so we don't append repeated data
                StringBuilder sb = new StringBuilder();
                while (reader.Peek() >= 0)
                {
                    // For each json line returned by baremetrics, dump to localFS, add to in mem cache
                    line = reader.ReadLine();
                    jsnf.WriteLine(line);
                    sb.AppendLine(line);
                }
                jsnf.Close();
                data.Close();
                reader.Close();
                UpdateRTD("baremetrics", qkey, "status", "complete");
                UpdateRTD("baremetrics", "all", "count", String.Format("{0}", m_BareCount++));
                Logr.Log(String.Format("baremetrics qkey({0}) complete count({1})", qkey, lineCount));
                // BareMetricsSummary summary = JsonConvert.DeserializeObject<BareMetricsSummary>(sb.ToString());
                dynamic updates = JsonConvert.DeserializeObject( sb.ToString( ) );
                s_Cache.UpdateBareCache(qkey, updates);
                UpdateRTD("baremetrics", qkey, "count", String.Format("{0}", updates.metrics.Count));
                return true;
            }
            catch (System.IO.IOException ex)
            {
                Logr.Log(String.Format("baremetrics qkey({0}) url({1}) {2}", qkey, url, ex));
            }
            catch (System.Net.WebException ex)
            {
                Logr.Log(String.Format("baremetrics qkey({0}) url({1}) {2}", qkey, url, ex));
            }
            return false;
        }
        #endregion
    }
}
