using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WebSocket4Net;

namespace SSAddin {
    class TiingoRealTimeMessageHandler {

        public delegate void RTDUpdate( string key, string subelem, string val);
        public delegate void SetSubID( string subID);

        protected static DataCache  s_Cache = DataCache.Instance( );

        protected RTDUpdate m_RTDUpdate;
        protected SetSubID  m_SetSubID;
        protected WebSocket m_Socket;
        protected int       m_HBCount;      // tiingo websock heartbeat count
        protected string    m_DefaultKey;   // key for non ticker specific updates

        #region Worker thread

        public TiingoRealTimeMessageHandler( WebSocket ws, RTDUpdate u, string dkey, SetSubID ssid) {
            // RTDServer.GetInstance( ) can instance the RTD server and all the RTD
            // COM machinery the first time we call it so we won't init as a static.
            // TiingoRealTimeMessageHandler and TWSCallback only get instanced when 
            // a worksheet has called twebsock, so we know it's OK to instance 
            // RTDServer at this point as the user definitely wants to do RT stuff.
            m_RTDUpdate = u;
            m_Socket = ws;
            m_HBCount = 0;
            m_DefaultKey = dkey;
            m_SetSubID = ssid;
        }

        public void MessageReceived(IDictionary<string, object> msg) {
            if (msg == null) {
                Logr.Log(String.Format("TiingoRealTimeMessageHandler.MessageReceived: null msg!"));
                return;
            }
            if (!msg.ContainsKey("messageType")) {
                Logr.Log(String.Format("TiingoRealTimeMessageHandler.MessageReceived: missing messageType field!"));
                return;
            }
            // for example messages https://api.tiingo.com/docs/iex/realtime#priceData
            string mt = msg["messageType"].ToString( );
            switch (mt) {
                case "I":   // Informational
                    // TODO: is there a subID?
                    if (msg.ContainsKey( "data" )) {
                        var dd = (IDictionary<string, object>)msg["data"];
                        if (dd.ContainsKey( "subscriptionId" )) {
                            m_SetSubID( dd["subscriptionId"].ToString( ) );
                        }
                    }
                    break;
                case "H":   // Heartbeat
                    m_HBCount++;
                    m_RTDUpdate(m_DefaultKey, "hbcount", m_HBCount.ToString());
                    break;
                case "A":   // Market data
                    break;
                case "E":   // Error
                    break;
                default:
                    Logr.Log( String.Format("TiingoRealTimeMessageHandler.MessageReceived: unexpected messageType({0})!", mt));
                    break;
            }
        }

        #endregion Worker thread

        #region Pool thread



        #endregion Pool thread

    }
}
