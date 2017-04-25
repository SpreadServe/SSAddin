Installing The SpreadServe Addin
================================

**Installing the addin for the first time**

* Get the XLL from http://spreadserve.com/s3/downloads.html or source from https://github.com/SpreadServe/SSAddin
* Install SSAddin.xll as an Excel addin.
  
  * Use SSAddin64.xll if you're running a 64 bit Excel.
  * Watch this video if you're unsure about adding an addin https://www.youtube.com/watch?v=i_sijj1NZFM
  * Put SSAddin.xll.config in the same directory as SSAddin.xll, and edit it to add Tiingo, Quandl and Baremetrics keys if you use those services.
  
* Create a new sheet, or load one of the test sheets to check that the addin is loaded.

  * Hit *fx* on the formula bar to get the Insert Function dialog.
  * Select the SpreadServe Addin function category.
  * You should see `s2cron`, `s2quandl` and other SpreadServe Addin functions listed.
  
* Bear in mind that the SpreadServe Addin does not add a ribbon menu. It's designed to work entirely
  through worksheet functions.

**SpreadServe Addin test sheets**

There are several test spreadsheets in the zip in the xls directory. These sheets all use RTD updates,
so make sure you are in automatic calculation mode. Go to ``Formulas/Calculation Options`` in Excel and
select Automatic. Use ctrl-alt-F9 to recalc everything and force the RTD subscriptions through.

* ``cron1.xls``: demonstrates the use of the ``s2cron`` and ``s2sub`` functions to set up and track a timer
  that goes off every 20 seconds. The timer will stop at the end of the day.
* ``cron2.xls``: uses of the ``s2cron`` and ``s2sub`` functions to set up and track a timer
  that goes off every 5 seconds. Note how the start and end dates are set in the s2cfg sheet so the
  timer will run beyond the end of the day, for as many days as the sheet is running.
* ``cron3.xls``: uses of the ``s2cron`` and ``s2sub`` functions to set up and track a timer
  that goes off daily at 1430. Note how the start and end dates are set in the s2cfg sheet.
* ``quandl1.xls``: uses ``s2quandl`` to launch a quandl query on a background thread in the subs sheet,
  and ``s2cache`` to pull the query result set into cells on the data sheet. You may have to ctrl-alt-F9
  a second time to force ``s2qcache`` execution in the subs sheet.  
* ``quandl2.xls``: a variation on quandl1. The two differences are the offsetting of the result set in
  the ``data`` sheet, and the use of ``s2vqcache`` instead of ``s2qcache``. The offsetting allows result sets
  to appear anywhere in a sheet instead of being anchored to the top left cell. ``s2vqcache`` is a volatile
  version of ``s2qcache``. Use of the volatile function avoids the need for a second ctrl-alt-F9.
* ``quandl3.xls``: combines the cron and quandl features to implement a quandl query that is executed every
  30 seconds.
* ``wsock.xls``: uses the `s2websock` function to subscribe to updates from an automated sheet hosted
  by SpreadServe.
* ``tiingows1.xls``: uses the `s2twebsock` and `s2sub` functions to subscribe to live ticking IEX market data
  from Tiingo. NB you will need to put your Tiingo authorization token into the s2cfg sheet to connect to Tiingo,
  and you'll need to be permissioned for IEX data at Tiingo.
* ``tiingows_option1.cls``: using `s2twebsock` and `s2sub` to drive a Black Scholes option calc with ticking
  IEX market data.
  
Some of the example sheets have ``_proxy`` suffixed to the name. These alternate versions are designed to work
from behind an internet proxy. They have extra config sheet entries to configure username, password and proxy
connection details. If you're in a corporate environment you'll probably need to use these.
