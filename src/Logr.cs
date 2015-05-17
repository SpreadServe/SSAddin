using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace SSAddin {
    class Logr {
        public static void Log( string ln ) {
            // Add code here to add threadId, processId and timestamp
            // Why? https://msdn.microsoft.com/en-us/library/system.diagnostics.tracelistener.traceoutputoptions%28v=vs.100%29.aspx
            // Yes: TextWriterTraceListener WriteLine ignores the TraceOptions that add ThreadId, ProcessId etc
            // And I can't get FileLogTraceListener to add those either. Can I be arsed with all the TraceEvent crap necessary?
            // No! I'm writing multi thread coded, so I want my log lines to have threadIds without any explicit code on my part!
            // Is that too much to ask?
            string tstamp = DateTime.Now.ToString( "o");
            int tid = Thread.CurrentThread.ManagedThreadId;
            int pid = Process.GetCurrentProcess( ).Id;
            Trace.WriteLine( String.Format( "{0} {1} {2} {3}", tstamp, pid, tid, ln ));
        }
    }
}
