# SSAddin
SSAddin, the SpreadServe Addin is a conventional Excel XLL addin implemented in C#. It has no build or run time dependencies on the [SpreadServe](http://spreadserve.com>) server runtime, and can be used independently in a regular desktop Excel installation, or in [SpreadServe](http://spreadserve.com>) itself. SSAddin supports quandl, cron style scheduled execution, and web socket live updates. SSAddin is freely available under the Apache License 2.0

## Binaries
You can download ready to install 32 & 64 bit binaries from [SpreadServe's download page](http://spreadserve.com/s3/downloads.html).

## Google Analytics
SSAddin gives you access to Google Analytics Reporting v3 API via the s2ganalytics and s2gacache worksheet functions. See the ganalytics_sessions2.xls worksheet for an example.

## Baremetrics
SSAddin gives you access to the Baremetrics API via the s2baremetrics and s2bcache worksheet functions. See the baremetrics_summary1.xlsx worksheet for an example.

## Quandl
quandl.com already distributes a perfectly good Excel addin, so how is SSAddin different? SSAddin uses no VBA and no GUI. It has no menu cluttering your Excel menu bar, no dialog or message boxes popping up. Everything is achieved via worksheet functions, and all network round trips are handled on a background thread so your Excel UI never blocks waiting for data to download from quandl.com.

## Tiingo
Tiingo is an exciting new financial data portal challenging high priced incumbents like Bloomberg and Thomson Reuters. Recently tiingo.com has added API access to historical data, which is now supported by SSAddin. 

## Cron
SSAddin enables the creation of cron style timer jobs in Excel that trigger RTD updates on the schedule you specify. Cron timers can be used to trigger recalculations, or to launch scheduled downloads from quandl.com

## Today
SSAddin provides the s2today function. s2today is a non volatile eqivalent of Excel's TODAY. It enables invocation in spreadsheets using RTD without triggering endless calc cycles. 

## Web sockets
SSAddin supports subscription to live ticking web data via web sockets.

## Documentation
http://spreadserve-addin.readthedocs.org/en/latest/index.html

## Acknowledgements
SSAddin builds on several other excellent OSS projects: Excel-DNA, NCrontab, WebSockets4Net and JSON.NET. We use a modified NCrontab that extends Unix style cron schedules to allow more finegrained timing specifications using seconds as well as minutes, hours, days and days of the week.

## Contact
john dot osullivan at spreadserve dot com

