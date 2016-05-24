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

    class TWSCallback {
        public delegate void ClosedCB( string wskey );

        protected WebSocket m_Client;
        protected ClosedCB m_ClosedCB;
        protected String m_AuthToken;
        protected String m_Key;
        protected String m_URL;
        protected String m_ProxyHost;
        protected String m_ProxyPort;
        protected String m_ProxyUser;
        protected String m_ProxyPassword;
        protected String m_SubscribeMessage;
        protected Dictionary<String, SortedSet<String>> m_Subscriptions = new Dictionary<string,SortedSet<string>>( );

        protected static DataCache s_Cache = DataCache.Instance( );
        // Braces are special chars in C# format strings, so we need a double brace to indicate a literal single brace,
        // rather than a the start of a place holder eg {0}
        protected static String s_SubscribeMessageFormat = "{{\"eventName\":\"subscribe\",\"eventData\":{{\"authToken\": \"{0}\"}}}}";

        protected static JsonSerializerSettings s_JsonSettings = new JsonSerializerSettings( )
        {
            Converters = { new JsonToDictionary( ) }
        };

        #region Worker thread

        public TWSCallback( Dictionary<string,string> work, ClosedCB ccb ) {
            // string host = 
            // string port = 
            // IPHostEntry he = Dns.GetHostEntry( host);
            // var proxy = new HttpConnectProxy( new IPEndPoint( IPAddressList[0], port));
            // m_Client.Proxy = ( SuperSocket.ClientEngine.IProxyConnector)proxy;
            m_Key = work["key"];
            m_URL = work["url"];
            work.TryGetValue( "auth_token", out m_AuthToken);
            m_ClosedCB = ccb;
            try {
                m_SubscribeMessage = String.Format( s_SubscribeMessageFormat, m_AuthToken );
                m_Client = new WebSocket( m_URL);
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

        void Subscribe( String ticker, String subelem) {
            SortedSet<String> ss = null;
            if (!m_Subscriptions.ContainsKey( ticker )) {
                m_Subscriptions[ticker] = new SortedSet<string>( );
            }
            ss = m_Subscriptions[ticker];
            ss.Add( subelem );
        }

        #endregion Worker thread

        #region Pool thread

        // All the pool thread methods are callbacks that will be fire on
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

        protected void HandleToken( JsonToken t, object v ) {
            // add code to stack token
        }

        protected void HandleValue( object v ) {
            // compare current stack to 
        }

        void MessageReceived( object sender, MessageReceivedEventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.MessageReceived: {0}", e.Message ) );
            var result = JsonConvert.DeserializeObject<IDictionary<string, object>>( e.Message, s_JsonSettings );
            RTDServer rtd = RTDServer.GetInstance( );
            if (rtd != null) {
                // rtd.CacheUpdateBatch( String.Format( "twebsock.{0}", m_Key), updates);
            }
            // s_Cache.UpdateWSCache( m_Key, updates );
        }

        void Closed( object sender, EventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.Closed: wskey({0})", m_Key ) );
            if (m_ClosedCB != null)
                m_ClosedCB( String.Format( "twebsock.{1}", m_Key));
            UpdateRTD( m_Key, "status", "closed" );
        }

        void Error( object sender, SuperSocket.ClientEngine.ErrorEventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.Error: wskey({0}) {1}", m_Key, e.Exception.Message ) );
            UpdateRTD( m_Key, "status", "error" );
            UpdateRTD( m_Key, "error", e.Exception.Message );
        }

        void Opened( object sender, EventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.Opened: subscribe message({0})", m_SubscribeMessage ) );
            m_Client.Send( m_SubscribeMessage );
            UpdateRTD( m_Key, "status", "open" );
        }

        #endregion Pool thread
    }
}
