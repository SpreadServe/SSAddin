# SSAddin
SSAddin, the SpreadServe Addin is a conventional Excel XLL addin implemented in C#. It has no build or run time dependencies on the `SpreadServe <http://spreadserve.com`_ server runtime, and can be used independently in a regular desktop Excel installation, or in `SpreadServe <http://spreadserve.com`_ itself. SSAddin supports quandl, cron style scheduled execution, and web socket live updates. SSAddin is freely available under the Apache License 2.0

## Acknowledgements
SSAddin builds on several other excellent OSS projects: Excel-DNA, NCrontab, WebSockets4Net and JSON.NET. We use a modified NCrontab that extends Unix style cron schedules to allow more finegrained timing specifications using seconds as well as minutes, hours, days and days of the week.

## Quandl
quandl.com already distributes a perfectly good Excel addin, so how is SSAddin different? SSAddin uses no VBA and no GUI. It has no menu cluttering your Excel menu bar, no dialog or message boxes popping up. Everything is achieved via worksheet functions, and all network round trips are handled on a background thread so your Excel UI never blocks waiting for data to download from quandl.com.

## Cron
SSAddin enables the creation of cron style timer jobs in Excel that trigger RTD updates on the schedule you specify. Cron timers can be used to trigger recalculations, or to launch scheduled downloads from quandl.com

## Web sockets
SSAddin supports subscription to live ticking web data via web sockets.

## Documentation
http://spreadserve-addin.readthedocs.org/en/latest/index.html