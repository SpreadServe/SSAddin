using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WebSocket4Net;
using Newtonsoft.Json;

namespace SSAddin {
    #region Deserialization classes
    // rtwebsvr websock format
    class SSWebCell {
        public String id { get; set; }
        public String body { get; set; }
    }
    #endregion

    class WSCallback {
        public delegate void ClosedCB( string wskey );

        protected string m_Key;
        protected WebSocket m_Client;
        protected ClosedCB m_ClosedCB;
        protected static DataCache s_Cache = DataCache.Instance( );

        #region Worker thread

        public WSCallback( Dictionary<string,string> work, ClosedCB ccb ) {
            // ctor will be invoked on the SSWebClient worker thread
            m_ClosedCB = ccb;
            m_Key = work["key"];
            string url = work["url"];
            try {
                m_Client = new WebSocket( url );
                m_Client.Opened += new EventHandler( Opened );
                m_Client.Error += new EventHandler<SuperSocket.ClientEngine.ErrorEventArgs>( Error );
                m_Client.Closed += new EventHandler( Closed );
                m_Client.MessageReceived += new EventHandler<MessageReceivedEventArgs>( MessageReceived );
                m_Client.DataReceived += new EventHandler<WebSocket4Net.DataReceivedEventArgs>( DataReceived );
                m_Client.Open( );
            }
            catch (System.ArgumentException ae) {
                Logr.Log( String.Format( "WSCallback.ctor: {0}", ae.Message ) );
            }
        }

        #endregion Worker thread

        #region Pool thread

        // All the pool thread methods are callbacks that will be fire on
        // web socket events on pool threads.

        protected void UpdateRTD( string qkey, string subelem, string value ) {
            // The RTD server doesn't necessarily exist. If no cell calls 
            // s2sub( ) it won't be instanced by Excel.
            RTDServer rtd = RTDServer.GetInstance( );
            if (rtd == null)
                return;
            string stopic = String.Format( "websock.{0}.{1}", qkey, subelem );
            rtd.CacheUpdate( stopic, value );
        }

        void DataReceived( object sender, WebSocket4Net.DataReceivedEventArgs e ) {
            Logr.Log( String.Format( "DataReceived: {0}", e.Data ) );
        }

        void MessageReceived( object sender, MessageReceivedEventArgs e ) {
            List<SSWebCell> updates = JsonConvert.DeserializeObject<List<SSWebCell>>( e.Message);
            Logr.Log( String.Format( "MessageReceived: updates.Count({0})", updates.Count ) );
            if (updates.Count == 0)
                return;
            RTDServer rtd = RTDServer.GetInstance( );
            if (rtd != null) {
                rtd.CacheUpdateBatch( String.Format( "websock.{0}", m_Key), updates);
            }
            s_Cache.UpdateWSCache( m_Key, updates );
        }

        void Closed( object sender, EventArgs e ) {
            Logr.Log( String.Format( "Closed: wskey({0})", m_Key ) );
            if (m_ClosedCB != null)
                m_ClosedCB( m_Key );
            UpdateRTD( m_Key, "status", "closed" );
        }

        void Error( object sender, SuperSocket.ClientEngine.ErrorEventArgs e ) {
            Logr.Log( String.Format( "Error: wskey({0}) {1}", m_Key, e.Exception.Message ) );
            UpdateRTD( m_Key, "status", "error" );
            UpdateRTD( m_Key, "error", e.Exception.Message );
        }

        void Opened( object sender, EventArgs e ) {
            Logr.Log( String.Format( "Opened: wskey({0})", m_Key ) );
            UpdateRTD( m_Key, "status", "open" );
        }

        #endregion Pool thread
    }
}
