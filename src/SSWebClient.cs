using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Net;
using System.Diagnostics;
using NCrontab;


namespace SSAddin {
    class SSWebClient {
        protected static DataCache      s_Cache = DataCache.Instance( );
        protected static char[]         csvDelimiterChars = { ',' };
        protected static SSWebClient    s_Instance;
        protected static object         s_InstanceLock = new object( );

        protected Queue<String[]>   m_InputQueue;
        protected HashSet<String>   m_InFlight;
        protected Dictionary<String, WSCallback> m_WSCallbacks;

        protected Thread            m_WorkerThread;     // for executing the web query
        protected ManualResetEvent  m_Event;            // control worker thread sleep
        protected String            m_TempDir;


        #region Excel thread methods

        protected SSWebClient( )
        {
            m_TempDir = System.IO.Path.GetTempPath( );
            m_InputQueue = new Queue<String[]>( );
            m_InFlight = new HashSet<String>( );
            m_WSCallbacks = new Dictionary<string, WSCallback>( );
            m_Event = new ManualResetEvent( false);
            m_WorkerThread = new Thread( BackgroundWork );
            m_WorkerThread.Start( );
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


        public bool AddRequest( string type, string key, String url) {
            // Is this job pending or in progress?
            string fkey = String.Format( "{0}.{1}", type, key);
            lock (m_InFlight) {
                if (m_InFlight.Contains( fkey )) {   // Queued or running...
                    Logr.Log( String.Format( "AddRequest: {0} is inflight", fkey ) );
                    return false;                   // so bail
                }
                lock (m_InputQueue) {
                    // Running on the main Excel thread here. Q the work, and
                    // signal the background thread to wake up...
                    Logr.Log( String.Format( "~A AddRequest adding {0} {1} {2}", type, key, url ) );
                    String[] req = { type, key, url};
                    m_InputQueue.Enqueue( req);
                }
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

        #endregion

        #region Worker thread methods

        protected String[] GetWork( ) {
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
                String[] work = GetWork( );
                while ( work != null) {
                    if (work[0] == "stop") {
                        // exit req from excel thread
                        Logr.Log( String.Format( "~A BackgroundWork thread exiting" ) );
                        return;
                    }
                    string fkey = String.Format( "{0}.{1}", work[0], work[1]);
                    Logr.Log( String.Format( "~A BackgroundWork new request fkey({0})", fkey) );
                    if (work[0] == "quandl") {
                        bool ok = DoQuandlQuery( work[1], work[2]);
                        lock (m_InFlight) {
                            m_InFlight.Remove( fkey );
                        }
                    }
                    else if ( work[0] == "websock") {
                        WSCallback wscb = new WSCallback( work[1], work[2], this.WSCallbackClosed );
                        lock (m_InFlight) {
                            m_WSCallbacks.Add( fkey, wscb );
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

        protected void UpdateRTD( string qkey, string subelem, string value ) {
            // The RTD server doesn't necessarily exist. If no cell calls 
            // s2sub( ) it won't be instanced by Excel.
            RTDServer rtd = RTDServer.GetInstance( );
            if ( rtd == null)
                return;
            string stopic = String.Format( "quandl.{0}.{1}", qkey, subelem );
            rtd.CacheUpdate( stopic, value );
        }


        protected bool DoQuandlQuery( string qkey, string url )
		{
			try	{
                string line = "";
                string lineCount = "0";
                // Set up the web client to HTTP GET
                var client = new WebClient( );
                Stream data = client.OpenRead( url);
                var reader = new StreamReader( data);
                // Local file to dump result
                int pid = Process.GetCurrentProcess( ).Id;
                string csvfname = String.Format( "{0}\\{1}_{2}.csv", m_TempDir, qkey, pid );
                Logr.Log( String.Format( "running quandl qkey({0}) {1} persisted at {2}", qkey, url, csvfname));
                var csvf = new StreamWriter( csvfname );
                UpdateRTD( qkey, "status", "starting" );
                // Clear any previous result from the cache so we don't append repeated data
                s_Cache.ClearQuandl( qkey );
                while ( reader.Peek( ) >= 0) {
                    // For each CSV line returned by quandl, dump to localFS, add to in mem cache, and 
                    // send a line count update to any RTD subscriber
                    line = reader.ReadLine( );
                    csvf.WriteLine( line );
                    lineCount = String.Format( "{0}", s_Cache.AddQuandlLine( qkey, line.Split( csvDelimiterChars)));
                    UpdateRTD( qkey, "count", lineCount );
                }
                csvf.Close( );
                data.Close();
                reader.Close();
                UpdateRTD( qkey, "status", "complete" );
                // TODO: add code here to report query status via RTD
                Logr.Log( String.Format( "quandl qkey({0}) complete count({1})", qkey, lineCount));
                return true;
			}
			catch( System.IO.IOException ex) {
                Logr.Log( String.Format( "quandl qkey({0}) {1}", qkey, ex) );
			}
            catch (System.Net.WebException ex) {
                Logr.Log( String.Format( "quandl qkey({0}) {1}", qkey, ex ) );
            }
            return false;
		}
        #endregion
    }
}
