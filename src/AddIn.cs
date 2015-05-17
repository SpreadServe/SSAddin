// Copyright Babbington Slade Ltd
using System;
using System.Diagnostics;
using ExcelDna.Integration;

namespace SSAddin {
	public class AddIn : IExcelAddIn {
        private static TraceSource m_TraceSource = new TraceSource( "ssaddin");

		public void AutoOpen( ) {
            Logr.Log( "AutoOpen" );
			ExcelIntegration.RegisterUnhandledExceptionHandler(e => "ERROR: " + (e as Exception).Message);
			//var excel = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
			//var xllPath = (string)XlCall.Excel(XlCall.xlGetName);
			//excel.AddIns.Add(xllPath, false /* don't copy file */).Installed = true;
		}

		public void AutoClose( ) {
            Logr.Log( "AutoClose" );
		}
	}
}
