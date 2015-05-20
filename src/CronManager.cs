using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Text;
using System.Timers;
using NCrontab;

namespace SSAddin {
    class CronManager {
        protected static object s_InstanceLock = new object( );
        protected static CronManager s_Instance;

        protected class CronTimer {
            protected IEnumerable<DateTime> m_Schedule;
            protected IEnumerator<DateTime> m_Iterator;
            protected Timer m_Timer;
            protected String m_Key;
            protected int m_Count = 0;
            protected String m_LastEventTime = "";
            protected String m_NextEventTime = "";

            #region Excel thread

            public CronTimer( String ckey, IEnumerable<DateTime> s ) {
                m_Key = ckey;
                m_Schedule = s;
                m_Iterator = s.GetEnumerator( );
                m_Timer = new System.Timers.Timer( );
                m_Timer.Enabled = false;
                m_Timer.AutoReset = false;
                m_Timer.Elapsed += this.OnTimerEvent;
                ScheduleTimer( );
            }

            #endregion Excel thread

            #region Pool thread

            public bool ScheduleTimer( ) {
                // NB this method is called by the Excel thread in the first instance. Subsequent calls are
                // on a pool thread, as .Net dispatches timer callbacks on pool threads, unless we specify otherwise.
                // I'm not bothering with a lock because the first timer event isn't scheduled until we do
                // m_Timer.Enabled = true below, and we don't touch the iterator after that, so there's no
                // chance that we'll have two threads touching m_Iterator at the same time. 
 
                // What if the next time we got from m_Iterator is already in the past?
                // If it is keep moving fwd til we get a time in the future.
                while ( m_Iterator.Current.CompareTo( DateTime.Now ) <= 0) {
                    if ( !m_Iterator.MoveNext( )) {
                        Logr.Log( String.Format( "ScheduleTimer: {0} exhausted", m_Key ) );
                        return false;
                    }
                }
                m_LastEventTime = DateTime.Now.ToString( );
                m_NextEventTime = m_Iterator.Current.ToString( );
                // Ticks is number of 100 nanos since 0001-01-01T00:00:00. Diff between now
                // and next event time 10K is the number of millisecs until the next cron event
                // for ckey. https://msdn.microsoft.com/en-us/library/system.datetime.ticks%28v=vs.100%29.aspx
                long ticks = m_Iterator.Current.Ticks - DateTime.Now.Ticks;
                long interval = Math.Abs( ticks / 10000 );
                if (interval == 0) {
                    Logr.Log( String.Format( "ScheduleTimer: ZERO interval! ckey({0}) Current({1}) Now({2})", 
                                                                            m_Key, m_Iterator.Current, DateTime.Now ) );
                    return false;
                }
                m_Timer.Interval = interval;
                m_Timer.Enabled = true;
                Logr.Log( String.Format( "ScheduleTimer: ckey({0}) Current({1}) Now({2})", m_Key, m_Iterator.Current, DateTime.Now ) );
                return true;
            }

            protected void UpdateRTD( string qkey, string subelem, string value ) {
                // The RTD server doesn't necessarily exist. If no cell calls 
                // s2sub( ) it won't be instanced by Excel.
                RTDServer rtd = RTDServer.GetInstance( );
                if (rtd == null)
                    return;
                string stopic = String.Format( "cron.{0}.{1}", qkey, subelem );
                rtd.CacheUpdate( stopic, value );
            }

            protected void OnTimerEvent( object o, ElapsedEventArgs e ) {
                m_Count++;
                ScheduleTimer( );
                Logr.Log( String.Format( "OnTimerEvent count({0}) last({1}) next({2})",
                                            m_Count, m_LastEventTime, m_NextEventTime ));
                UpdateRTD( m_Key, "count", Convert.ToString( m_Count) );
                UpdateRTD( m_Key, "last", m_LastEventTime );
                UpdateRTD( m_Key, "next", m_NextEventTime );
            }
            #endregion Pool thread
        }

        protected Dictionary<string, CronTimer> m_CronMap = new Dictionary<string, CronTimer>( );

        #region Excel thread

        public static CronManager Instance( ) {
            // Unlikley that two threads will attempt to instance this singleton at the
            // same time, but we'll lock just in case.
            lock (s_InstanceLock) {
                if (s_Instance == null) {
                    s_Instance = new CronManager( );
                }
                return s_Instance;
            }
        }

        protected CronManager( ) {

        }

        public bool AddCron( string ckey, string cronex, DateTime start, DateTime end) {
            // no locking here as we won't touch any objects that are shared with another thread
            string[] cronflds = new string[6];
            Logr.Log( String.Format( "AddCron: cronex({0}) start({1}) end({2})", cronex, start, end) );

            try {
                CrontabSchedule schedule = CrontabSchedule.Parse( cronex );
                IEnumerable<DateTime> numerable = schedule.GetNextOccurrences( start, end );
                m_CronMap[ckey] = new CronTimer( ckey, numerable );
                return true;
            }
            catch (Exception ex) {
                Logr.Log( String.Format( "AddCron: {0}", ex.Message ) );
                return false;
            }
        }
        #endregion Excel thread
    }
}
