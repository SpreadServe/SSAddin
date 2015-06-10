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
            // Keep internal copies of ctor parms
            protected String m_Key;
            protected String m_Cronex;
            protected DateTime? m_Start;
            protected DateTime? m_End;
            // Working storage for the timer
            protected IEnumerable<DateTime> m_Schedule;
            protected IEnumerator<DateTime> m_Iterator;
            protected Timer m_Timer;
            protected int m_Count = 0;
            protected String m_LastEventTime = "";
            protected String m_NextEventTime = "";
            protected bool m_Closed = false;

            #region Excel thread

            // Only the Excel thread should touch m_Timer.Enabled. If a pool thread can flip the
            // state of Enabled we may get race conditions with timers being inadvertently
            // re-enabled after being switched off.

            public CronTimer( String ckey, String cronex, DateTime? start, DateTime? end ) {
                // Save setup parms so later invocations of AddCron can check whether we need
                // to create a new instance or not.
                m_Key = ckey;
                m_Cronex = cronex;
                m_Start = start;
                m_End = end;
                // Now do the real biz of setting up the timer.
                DateTime sStart = start ?? DateTime.Now;
                DateTime sEnd = end ?? new DateTime( sStart.Year, sStart.Month, sStart.Day, 23, 59, 59 );
                CrontabSchedule schedule = CrontabSchedule.Parse( cronex );
                m_Schedule = schedule.GetNextOccurrences( sStart, sEnd );
                m_Iterator = m_Schedule.GetEnumerator( );
                m_Timer = new System.Timers.Timer( );
                m_Timer.Enabled = false;
                m_Timer.AutoReset = false;
                m_Timer.Elapsed += this.OnTimerEvent;
                ScheduleTimer( );
                m_Timer.Enabled = true;
            }

            public void Close( ) {
                // There's a nasty bug in the Timer class: setting the interval schedules
                // a timer callback even if Enabled is set false. This means you can't
                // stop a timer by setting Enabled=false if there is another callback
                // scheduled, and that callback will set Interval!
                // https://evolpin.wordpress.com/2014/04/25/the-curious-case-of-system-timers-timer/
                // So we set a flag to tell the ScheduleTimer( ) method not to touch
                // m_Timer.Interval. And then hopefully, the GC will do it's stuff...
                m_Closed = true;
            }

            public string Cronex {
                get { return m_Cronex; }
            }

            public DateTime? Start {
                get { return m_Start; }
            }

            public DateTime? End {
                get { return m_End; }
            }

            #endregion Excel thread

            #region Pool thread

            public bool ScheduleTimer( ) {
                if (m_Closed) {
                    Logr.Log( String.Format( "ScheduleTimer: {0} is closed", m_Key ) );
                    return false;
                }
                // NB this method is called by the Excel thread in the first instance. Subsequent calls are
                // on a pool thread, as .Net dispatches timer callbacks on pool threads, unless we specify otherwise.
                // I'm not bothering with a lock because the first timer event isn't scheduled until we do
                // m_Timer.Enabled = true below, and we don't touch the iterator after that, so there's no
                // chance that we'll have two threads touching m_Iterator at the same time. 
 
                // What if the next time we got from m_Iterator is already in the past?
                // If it is keep moving fwd til we get a time in the future. Bear in mind there's
                // an error condition where Current can appear to be in the future when it's not.
                // If DateTime.Now==2015-06-03T20:03:04.9998700, and Current==2015-06-03T20:03:05
                // then CompareTo will tell us that Current is in the future, when for our
                // purposes it's not. If ticks is -ve then Current is in the past. If Current is
                // a small +/-ve then it's the same as Now since our unit of granularity in the 
                // cron sys is 1 sec. There are 10,000 ticks to the millisec, 10,000,000 to the sec.
                // So we'll look for Current to be 1,000,000 ticks later than Now before scheduling.
                // Which is +1,000,000. If ticks is -ve then Current is in the past. This check 
                // should also prevent interval==0 below!
                long ticks = m_Iterator.Current.Ticks - DateTime.Now.Ticks;
                while ( ticks < 1000000) {
                    if ( !m_Iterator.MoveNext( )) {
                        Logr.Log( String.Format( "ScheduleTimer: {0} exhausted", m_Key ) );
                        return false;
                    }
                    ticks = m_Iterator.Current.Ticks - DateTime.Now.Ticks;
                }
                m_LastEventTime = DateTime.Now.ToString( );
                m_NextEventTime = m_Iterator.Current.ToString( );
                // Ticks is number of 100 nanos since 0001-01-01T00:00:00. Diff between now
                // and next event time 10K is the number of millisecs until the next cron event
                // for ckey. https://msdn.microsoft.com/en-us/library/system.datetime.ticks%28v=vs.100%29.aspx
                // 10,000 ticks in 1 millisec
                long interval = Math.Abs( ticks / 10000 );
                if (interval == 0) {
                    // Given the code above, this should not happen!
                    Logr.Log( String.Format( "ScheduleTimer: ZERO interval! ckey({0}) Current({1}) Now({2})", 
                                                                            m_Key, m_Iterator.Current, DateTime.Now ) );
                    return false;
                }
                m_Timer.Interval = interval;
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

        public bool AddCron( string ckey, string cronex, DateTime? start, DateTime? end) {
            // no locking here as we won't touch any objects that are shared with another thread
            string[] cronflds = new string[6];
            Logr.Log( String.Format( "AddCron: cronex({0}) start({1}) end({2})", cronex, start, end) );
            try {
                if (m_CronMap.ContainsKey( ckey )) {
                    // If there's already a timer with ckey it may be that an Excel users has triggered
                    // another invocation of s2cron( ) by editting the s2cfg sheet, or with a sh-ctrl-alt-F9.
                    // Either way, we need to remove the old timer, and create a new one, but only if the
                    // new one is different.
                    CronTimer oldTimer = m_CronMap[ckey];
                    if (oldTimer.Cronex == cronex && oldTimer.Start == start && oldTimer.End == end) {
                        // no change, so we won't overwrite the entry for ckey
                        return true;
                    }
                    m_CronMap.Remove( ckey);
                    oldTimer.Close( );
                }
 
                m_CronMap[ckey] = new CronTimer( ckey, cronex, start, end );
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
