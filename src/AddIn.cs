// Copyright Babbington Slade Ltd
using System;
using System.Diagnostics;
using System.Net.PeerToPeer;
using ExcelDna.Integration;

namespace SSAddin {
	public class AddIn : IExcelAddIn {
        private static TraceSource m_TraceSource = new TraceSource( "ssaddin");

        // We don't use System.Net.PeerToPeer.PeerName anywhere in the SSAddin code.
        // This is here to force loading of System.Net.dll 4.0.0.0 before Google.Apis.dll
        // asks for v2.0.5.0 as described in the link below. For some reason assembly 
        // bindingRedirects in SSAddin.xll.config don't work for System.Net.dll, even
        // though they do work for System.Net.Http.Primitives. I don't want to have
        // to require the KB2468871 fix, which wouldn't install on my Win8 laptop
        // anyway. One more thing: if you're trying to do bindingRedirects in your
        // .xll.config, do make sure ExcelDnaPack isn't packing it into the XLL. It
        // will if it find a .xll.config with a matching base name next to the .dna
        // file. JOS 2017-05-06
        // https://github.com/google/google-api-dotnet-client/issues/378
        // http://blog.slaks.net/2013-12-25/redirecting-assembly-loads-at-runtime/
        private static PeerName m_PeerName = new PeerName( "dummy" );

		public void AutoOpen( ) {
            Logr.Log( "AutoOpen" );
			ExcelIntegration.RegisterUnhandledExceptionHandler( e => "EXCEPTION: " + (e as Exception).Message);
		}

		public void AutoClose( ) {
            Logr.Log( "AutoClose" );
		}
	}
}
