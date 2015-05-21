using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace SSAddin {
    class Logr {

        protected static TextWriterTraceListener m_TextOut;

        static Logr( ) {
            // Trace goes to /dev/null in Release mode unless some kind of debug tool is used.
            // Need to add a listener to send to localFS. JOS 2015-05-19
            string logpath = String.Format( "{0}\\ssaddin_{1}.log", System.IO.Path.GetTempPath( ), Process.GetCurrentProcess( ).Id);
            m_TextOut = new TextWriterTraceListener( System.IO.File.CreateText( logpath));
            Trace.Listeners.Add( m_TextOut);
        }

        public static void Log( string ln ) {
            // Add code here to add threadId, processId and timestamp
            // Why? https://msdn.microsoft.com/en-us/library/system.diagnostics.tracelistener.traceoutputoptions%28v=vs.100%29.aspx
            // Yes: TextWriterTraceListener WriteLine ignores the TraceOptions that add ThreadId, ProcessId etc
            // And I can't get FileLogTraceListener to add those either. Can I be arsed with all the TraceEvent crap necessary?
            // No! I'm writing multi thread code, so I want my log lines to have threadIds without any explicit code on my part!
            // Is that too much to ask?
            string tstamp = DateTime.Now.ToString( "o");
            int tid = Thread.CurrentThread.ManagedThreadId;
            int pid = Process.GetCurrentProcess( ).Id;
            Trace.WriteLine( String.Format( "{0} {1} {2} {3}", tstamp, pid, tid, ln ));
            // In release mode writes are buffered, but we want to see them logged immediately, so flush.
            Trace.Flush( );
        }
    }
}
