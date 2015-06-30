// Copyright Babbington Slade Ltd
using System;
using System.Diagnostics;
using ExcelDna.Integration;

namespace SSAddin {
	public class AddIn : IExcelAddIn {
        private static TraceSource m_TraceSource = new TraceSource( "ssaddin");

		public void AutoOpen( ) {
            Logr.Log( "AutoOpen" );
			ExcelIntegration.RegisterUnhandledExceptionHandler( e => "EXCEPTION: " + (e as Exception).Message);
		}

		public void AutoClose( ) {
            Logr.Log( "AutoClose" );
		}
	}
}
