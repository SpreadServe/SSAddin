using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using WebSocket4Net;
using Newtonsoft.Json;
using SuperSocket.ClientEngine;

namespace SSAddin {

    // Apart from the ctor every method in TWSCallback that reads/writes object state locks
    // because we may have the worker background thread in touching object state as well as
    // a pool thread firing by a socket callback.
    class TWSCallback {
        public delegate void ClosedCB( string wskey );

        protected WebSocket m_Client;
        protected ClosedCB  m_ClosedCB;
        protected String    m_AuthToken;
        protected String    m_Key;
        protected String    m_URL;
        protected String    m_ProxyHost;
        protected String    m_ProxyPort;
        protected String    m_ProxyUser;
        protected String    m_ProxyPassword;
        protected String    m_SubID;
        protected Dictionary<String, SortedSet<String>> m_Subscriptions = new Dictionary<string,SortedSet<string>>( );
        protected TiingoRealTimeMessageHandler m_RTMHandler;
        protected SortedSet<string> m_PendingSubs = new SortedSet<string>( );
        // m_MktDataRecord field names must match up with https://api.tiingo.com/docs/iex/realtime
        // type: Q for quote, T for trade
        protected string[] m_MktDataRecord = { "type", "timestamp", "tickid", "ticker", "bidsz", "bid", "mid", "ask", "asksz", "ltrade", "ltradesz"};
        protected int m_TickerIndex = 3;    // index of "ticker" in m_MktDataRecord

        protected static DataCache s_Cache = DataCache.Instance( );
        // Braces are special chars in C# format strings, so we need a double brace to indicate a literal single brace,
        // rather than a the start of a place holder eg {0}
        // This is the old s_SubscribeMessageFormat from before Rishi introduced subscriptionIDs
        // protected static String s_SubscribeMessageFormat = "{{\"eventName\":\"subscribe\",\"eventData\":{{\"authToken\": \"{0}\"}}}}";
        protected static String s_SubscribeMessageFormat = "{{ \"eventName\":\"subscribe\",\"authorization\":\"{0}\",\"eventData\":{{ {1} }} }}";
        // Format for composing {1} in s_SubscribeMessageFormat.
        protected static String s_EventDataFormat = "\"thresholdLevel\":0,\"tickers\":[{0}]";
        protected static String s_EventDataSubIdFormat = "\"subscriptionId\":{0},\"thresholdLevel\":0,\"tickers\":[{1}]";

        protected static JsonSerializerSettings s_JsonSettings = new JsonSerializerSettings( ) {
            Converters = { new JsonToDictionary( ) },
            NullValueHandling = NullValueHandling.Ignore
        };

        #region Worker thread

        public TWSCallback( Dictionary<string,string> work, ClosedCB ccb ) {
            // No need to bother locking in the ctor. We are on the background
            // worker thread here, but we won't get methods fired on the pool
            // threads until this method exits.
            m_Key = work["key"];
            m_URL = work["url"];
            work.TryGetValue( "auth_token", out m_AuthToken);
            m_ClosedCB = ccb;
            try {
                m_Client = new WebSocket( m_URL);
                m_RTMHandler = new TiingoRealTimeMessageHandler(m_Client, MktDataTick, HeartBeat, SetSubID);
                m_Client.Opened += new EventHandler( Opened );
                m_Client.Error += new EventHandler<SuperSocket.ClientEngine.ErrorEventArgs>( Error );
                m_Client.Closed += new EventHandler( Closed );
                m_Client.MessageReceived += new EventHandler<MessageReceivedEventArgs>( MessageReceived );
                m_Client.DataReceived += new EventHandler<WebSocket4Net.DataReceivedEventArgs>( DataReceived );
                // Do we need to set up a proxy?
                if (work.TryGetValue("http_proxy_host", out m_ProxyHost)) {
                    IPHostEntry he = Dns.GetHostEntry(m_ProxyHost);
                    int port = 80;
                    if (work.TryGetValue("http_proxy_host", out m_ProxyPort)) {
                        if (!Int32.TryParse(m_ProxyPort, out port))
                            port = 80;
                    }
                    var proxy = new HttpConnectProxy( new IPEndPoint( he.AddressList[0], port));
                    // Do we need to supply authentication to the proxy?
                    if (work.TryGetValue("http_proxy_user", out m_ProxyUser))
                    {
                        work.TryGetValue("http_proxy_password", out m_ProxyPassword);
                        // encode user:password as base64 and supply as 'Proxy-Authorization: Basic dXNlbWU6dGVzdA=='
                        string upass = String.Format("{0}:{1}", m_ProxyUser, m_ProxyPassword);
                        var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(upass);
                        string b64 = System.Convert.ToBase64String(plainTextBytes);
                        proxy.Authorization = String.Format("Basic {0}", b64);
                    }
                    m_Client.Proxy = proxy;
                }
                Logr.Log(String.Format("TWSCallback.ctor: key({0}) url({1}) auth_token({2}) http_proxy_host({3}) http_proxy_port({4}) http_proxy_user({5}) http_proxy_password({6})",
                                            m_Key, m_URL, m_AuthToken, m_ProxyHost, m_ProxyPort, m_ProxyUser, m_ProxyPassword));
                m_Client.Open( );
            }
            catch (System.ArgumentException ae) {
                Logr.Log( String.Format( "TWSCallback.ctor: {0}", ae.Message ) );
            }
            catch (System.FormatException fe) {
                Logr.Log( String.Format( "TWSCallback.ctor: format error fmt({0}) auth({1}) err({2})", s_SubscribeMessageFormat, m_AuthToken, fe.Message ) );
            }
        }

        public void AddSubscriptions( List<Dictionary<string,string>> subs)
        {
            // AddSubscriptions will be called from the background worker thread, but all the
            // object state it touches here could also be touched by the pool thread methods
            // below, so we lock on the socket object.
            lock (m_Client) {
                // First, put each subscription into the {ticker:[sub1,sub2] map
                // that tracks all existing ticker_sub subscriptions.
                // We're looking for eg twebsock.iex.appl_mid and we want to ignore
                // twebsock.iex.hbount.
                foreach (var sub in subs) {
                    if ( sub.ContainsKey("cachekey") && sub.ContainsKey("ticker_field")) {
                        string ckey = sub["cachekey"];
                        string tfld = sub["ticker_field"];
                        string[] tflds = tfld.Split('_');
                        if (ckey == m_Key && tflds.Length == 2) {
                            string ticker = tflds[0];
                            string subelem = tflds[1];
                            SortedSet<String> ss = null;
                            if (!m_Subscriptions.ContainsKey( ticker )) {
                                m_Subscriptions[ticker] = new SortedSet<string>( );
                                m_PendingSubs.Add( ticker );
                            }
                            ss = m_Subscriptions[ticker];
                            ss.Add( subelem );
                        }
                    }
                }
            }
            // A worksheet invocation of s2sub could cause the background thread to dispatch here at
            // any time. Maybe because the user has edited a sheet to add a new s2sub( ). Any new
            // subs that have accumulated as a result should be dispatched as eearly as possible.
            DispatchSubscriptions( );
        }

        #endregion Worker thread

        #region Worker or pool thread
        void DispatchSubscriptions( ) {
            // We could be called by either thread, so lock as we'll potentially change object state
            // and send stuff down the socker while another thread wants to do the same.
            lock (m_Client) {
                if (m_Client.State == WebSocketState.Open) {
                    // The socket is open, so the Opened( ) callback below must have already fired.
                    // An open socket isn't enough. We also need a subID, which we only have after
                    // we've processed the response to the initial subscription.
                    StringBuilder sb = new StringBuilder( );
                    int inx = 0;
                    foreach (string sub in m_PendingSubs) {
                        if (inx > 0)
                            sb.Append(",");
                        sb.Append( String.Format( "\"{0}\"", sub ) );
                        inx++;
                    }
                    string sublist = sb.ToString( );
                    string ed;
                    if (m_SubID != null) {
                        // We've got a subID, so only send a message if we've got tickers to add.
                        if (sublist.Length == 0)
                            return;
                        ed = String.Format(s_EventDataSubIdFormat, m_SubID, sublist);
                    }
                    else {
                        // We don't have a subID, so compose an initial message and send whether or not
                        // we have a ticker list.
                        ed = String.Format(s_EventDataFormat, sublist);
                    }
                    string submsg = String.Format( s_SubscribeMessageFormat, m_AuthToken, ed );
                    Logr.Log( String.Format( "TWSCallback.DispatchSubscriptions: subscribe message({0})", submsg ) );
                    m_Client.Send( submsg );
                    m_PendingSubs.Clear();
                }
            }
        }
        #endregion Worker or pool thread

        #region Pool thread

        // All the pool thread methods are callbacks that will be fired on
        // web socket events on pool threads.

        protected void UpdateRTD( string key, string subelem, string value ) {
            // The RTD server doesn't necessarily exist. If no cell calls 
            // s2sub( ) it won't be instanced by Excel.
            RTDServer rtd = RTDServer.GetInstance( );
            if (rtd == null)
                return;
            string stopic = String.Format( "twebsock.{0}.{1}", key, subelem );
            rtd.CacheUpdate( stopic, value );
        }

        void DataReceived( object sender, WebSocket4Net.DataReceivedEventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.DataReceived: {0}", e.Data ) );
        }

        void MessageReceived( object sender, MessageReceivedEventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.MessageReceived: {0}", e.Message ) );
            var msg = JsonConvert.DeserializeObject<IDictionary<string, object>>( e.Message, s_JsonSettings );
            m_RTMHandler.MessageReceived(msg);  // may cause callbacks to HeartBeat or MktDataTick below
        }

        public void HeartBeat(int hb) {
            UpdateRTD( m_Key, "hbcount", hb.ToString( ));
        }

        public void MktDataTick( IList<object> tick)
        {
            if (tick.Count < m_MktDataRecord.Length) {
                Logr.Log(String.Format("TWSCallback.MktDataTick: fld count under {0}! {1}", 
                                                        m_MktDataRecord.Length, tick.ToString( )));
                return;
            }
            string ticker = tick[m_TickerIndex].ToString();
            lock (m_Client) {
                if (!m_Subscriptions.ContainsKey(ticker))
                    return;
                SortedSet<string> fldset = m_Subscriptions[ticker];
                for ( int inx = 0; inx < m_MktDataRecord.Length; inx++) {
                    string fld = m_MktDataRecord[inx];
                    object val = tick[inx];
                    if ( fldset.Contains( fld) && val != null) {
                         UpdateRTD( m_Key, String.Format( "{0}_{1}", ticker, fld), tick[inx].ToString( ));
                    }
                }
            }
            // TODO: add TWSCache to s_Cache so that we only need an RTD sub to one field in a record, for
            // instance bid, and then the rest can be pulled from the cache...
            // s_Cache.UpdateWSCache( m_Key, updates );
        }

        void Closed( object sender, EventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.Closed: wskey({0})", m_Key ) );
            if (m_ClosedCB != null)
                m_ClosedCB( String.Format( "twebsock.{0}", m_Key));
            lock (m_Client) {
                // Socket has closed, and will need to be reopened. The reopen will trigger
                // another initial subscription, and then a new sub ID.
                m_SubID = null;
            }
            UpdateRTD( m_Key, "status", "closed" );
        }

        void Error( object sender, SuperSocket.ClientEngine.ErrorEventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.Error: wskey({0}) {1}", m_Key, e.Exception.Message ) );
            UpdateRTD( m_Key, "status", "error" );
            UpdateRTD( m_Key, "error", e.Exception.Message );
        }

        void Opened( object sender, EventArgs e ) {
            DispatchSubscriptions( );
            UpdateRTD( m_Key, "status", "open" );
        }

        void SetSubID( string subID ) {
            // This method gets called on a pool thread by TiingoRealTimeMessageHandler
            lock (m_Client) {
                m_SubID = subID;
            }
            DispatchSubscriptions( );
        }

        #endregion Pool thread
    }
}
