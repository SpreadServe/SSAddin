using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WebSocket4Net;
using Newtonsoft.Json;

namespace SSAddin {

    class TWSCallback {
        public delegate void ClosedCB( string wskey );

        protected WebSocket m_Client;
        protected ClosedCB m_ClosedCB;
        protected String m_AuthToken;
        protected String m_Key = "twebsock";
        protected String m_URL;
        protected String m_SubscribeMessage;
        protected Dictionary<String, SortedSet<String>> m_Subscriptions = new Dictionary<string,SortedSet<string>>( );

        protected static DataCache s_Cache = DataCache.Instance( );
        // Braces are special chars in C# format strings, so we need a double brace to indicate a literal single brace,
        // rather than a the start of a place holder eg {0}
        protected static String s_SubscribeMessageFormat = "{{'eventName':'subscribe','eventData':{{'authToken': '{0}'}}}}";

        #region Worker thread

        public TWSCallback( String url, String auth, ClosedCB ccb ) {
            m_AuthToken = auth;
            m_ClosedCB = ccb;
            m_URL = url;
            try {
                m_SubscribeMessage = String.Format( s_SubscribeMessageFormat, m_AuthToken );
                m_Client = new WebSocket( url );
                m_Client.Opened += new EventHandler( Opened );
                m_Client.Error += new EventHandler<SuperSocket.ClientEngine.ErrorEventArgs>( Error );
                m_Client.Closed += new EventHandler( Closed );
                m_Client.MessageReceived += new EventHandler<MessageReceivedEventArgs>( MessageReceived );
                m_Client.DataReceived += new EventHandler<WebSocket4Net.DataReceivedEventArgs>( DataReceived );
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

        protected void UpdateRTD( string ticker, string subelem, string value ) {
            // The RTD server doesn't necessarily exist. If no cell calls 
            // s2sub( ) it won't be instanced by Excel.
            RTDServer rtd = RTDServer.GetInstance( );
            if (rtd == null)
                return;
            string stopic = String.Format( "twebsock.{0}.{1}", ticker, subelem );
            rtd.CacheUpdate( stopic, value );
        }

        void DataReceived( object sender, WebSocket4Net.DataReceivedEventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.DataReceived: {0}", e.Data ) );
        }

        void MessageReceived( object sender, MessageReceivedEventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.MessageReceived: {0}", e.Message ) );

            RTDServer rtd = RTDServer.GetInstance( );
            if (rtd != null) {
                // rtd.CacheUpdateBatch( String.Format( "twebsock.{0}", m_Key), updates);
            }
            // s_Cache.UpdateWSCache( m_Key, updates );
        }

        void Closed( object sender, EventArgs e ) {
            Logr.Log( String.Format( "TWSCallback.Closed: wskey({0})", m_Key ) );
            if (m_ClosedCB != null)
                m_ClosedCB( String.Format( "{0}.{1}", m_Key, m_URL ));
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
