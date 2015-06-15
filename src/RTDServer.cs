using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace SSAddin {
    [ComVisible(true)]
    class RTDServer : ExcelRtdServer {

        // Use a static collection of instance refs as we can't use a singleton pattern here.
        // Excel fires the ctor for an RTDServer via COM. We'll track the instances, and
        // hand out refs to other classes as necessary. We shouldn't need to lock on this
        // object as all access should be on the Excel thread.
        static protected List<RTDServer>        s_Instances = new List<RTDServer>( );

        protected Dictionary<Topic,string>      m_Subscriptions;    // only Excel thread should touch this
        protected Dictionary<string, Topic>     m_TopicMap;         // Excel and worker thread touch this

        #region Excel thread

        public static RTDServer GetInstance( ) {
            // This should only be called by methods executing on the Excel thread as there's
            // no locking on s_Instances.
            if (s_Instances.Count == 0)
                return null;
            return s_Instances[0];
        }

        public RTDServer( ) {
            Logr.Log( "~A RTDServer created" );
            m_Subscriptions = new Dictionary<Topic, string>( );
            m_TopicMap = new Dictionary<string, Topic>( );
            s_Instances.Add( this );
        }

        protected int GetTopicId( Topic topic ) {
            // Why doesn't the Topic class expose the topicId ?
            var topicType = typeof( Topic );
            var topicId = topicType.GetField( "TopicId", BindingFlags.GetField | BindingFlags.Instance | BindingFlags.NonPublic );
            int rv = (int)topicId.GetValue( topic );
            return rv;
        }

        protected override bool ServerStart( ) {
            Logr.Log( "~A RTDServer.ServerStart" );
            return true;
        }

        protected override void ServerTerminate( ) {
            Logr.Log( "~A RTDServer.ServerTerminate" );
            // Clear down any running timers...
            CronManager.Instance( ).Clear( );
            s_Instances.Remove( this );
        }

        protected override object ConnectData( Topic topic, System.Collections.Generic.IList<string> topicInfo, ref bool newValues ) {
            lock (m_TopicMap) {
                string stopic = topicInfo[0];
                int topicId = GetTopicId( topic );
                Logr.Log( String.Format( "~A ConnectData: {0} - {1}", topicId, stopic ) );
                m_Subscriptions.Add( topic, stopic );
                m_TopicMap.Add( stopic, topic );
                return ExcelErrorUtil.ToComError( ExcelError.ExcelErrorNA );
            }
        }

        protected override void DisconnectData( Topic topic ) {
            lock (m_TopicMap) {
                string stopic = m_Subscriptions[topic];
                Logr.Log( String.Format( "~A DisconnectData: {0}", stopic ) );
                m_Subscriptions.Remove( topic );
                m_TopicMap.Remove( stopic );
            }
        }

        #endregion

        #region Worker or pool thread

        public void CacheUpdate( string stopic, string value ) {
            lock (m_TopicMap) {
                Logr.Log( String.Format( "~A CacheUpdate topic({0}) val({1})", stopic, value ) );
                if (m_TopicMap.ContainsKey( stopic )) {
                    Topic topic = m_TopicMap[stopic];
                    topic.UpdateValue( value );
                }
                else {
                    Logr.Log( String.Format( "~A CacheUpdate UNKNOWN topic({0}) val({1})", stopic, value ) );
                }
            }
        }

        public void CacheUpdateBatch( string stroot, List<SSWebCell> updates) {
            lock (m_TopicMap) {
                Logr.Log( String.Format( "~A CacheUpdateBatch {0} {1}", stroot, updates.Count));
                foreach ( SSWebCell wc in updates) {
                    String stopic = String.Format( "{0}.{1}", stroot, wc.id);
                    if (m_TopicMap.ContainsKey( stopic )) {
                        Topic topic = m_TopicMap[stopic];
                        topic.UpdateValue( wc.body);
                        Logr.Log( String.Format( "RTDServer.CacheUpdateBatch: topic({0}) value({1})", stopic, wc.body));
                    }
                }
            }
        }

        #endregion
    }
}
