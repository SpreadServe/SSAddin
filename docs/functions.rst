SpreadServe Addin Worksheet Functions
=====================================

These are the functions you can invoke directly from cells in your spreadsheet.

**s2about**: get version information.

Parameters

* None

Return value: a string detailing the SpreadServe Addin version, and the version of Excel hosting the adding.

**s2cron**: setup scheduled timer.

Parameters

* ``CronKey``: a value or cell reference evaluating to a string that matches a value in column C of
  the s2cfg sheet. The s2cfg row with the matching column C value will be used to specify a cron job.
  See the cron1, 2 or 3 example sheets.
  
Return value: "OK" if the function succeeds, an Excel error otherwise.

**s2quandl**: launch a quandl query.

Parameters

* ``QueryKey``: a value or cell reference evaluating to a string that matches a value in column C of
  the s2cfg sheet. The s2cfg row with the matching column C value will be used to specify a quandl query.
  See the quandl1, 2 or 3 example sheets.
* ``Trigger``: an optional trigger. The value isn't used inside the function, but a change in the input can
  be used to force repeat execution. See the quandl3.xls sheet for an example of an s2quandl trigger parameter
  hooked up to s2cron output to rerun a query on a timed basis.
  
Return value: "OK" if the function succeeds, an Excel error otherwise.

**s2qcache**: get a value from a quandl query result set. The position of the cell invoking this function is used
to figure out which cell to get from the result set.

Parameters

* ``QueryKey``: should match the QueryKey given to `s2quandl`.
* ``XOffset``: defaults to 0. If the left hand side of the result grid on your sheet is not column A this should
  be the number of columns across.
* ``YOffset``: defaults to 0. If the top row of the result grid on your sheet is not row 1 this should
  be the number of rows down.
* ``Trigger``: an optional trigger. The value isn't used inside the function, but a change in the input can
  be used to force repeat execution. See the quandl3.xls sheet for an example of an s2quandl trigger parameter
  hooked up to s2cron output to rerun a query on a timed basis.

Return value: a value from the result set, or #N/A.
  
**s2vqcache**: a volatile version of s2qcache.

Parameters

* ``QueryKey``: should match the QueryKey given to `s2quandl`.
* ``XOffset``: defaults to 0. If the left hand side of the result grid on your sheet is not column A this should
  be the number of columns across.
* ``YOffset``: defaults to 0. If the top row of the result grid on your sheet is not row 1 this should
  be the number of rows down.

Return value: a value from the result set, or #N/A.

**s2tiingo**: launch a tiingo query.

Parameters

* ``QueryKey``: a value or cell reference evaluating to a string that matches a value in column C of
  the s2cfg sheet. The s2cfg row with the matching column C value will be used to specify a tiingo query.
  See the tiingo1 or 2 example sheets.
* ``Trigger``: an optional trigger. The value isn't used inside the function, but a change in the input can
  be used to force repeat execution. 
  
Return value: "OK" if the function succeeds, an Excel error otherwise.

**s2tcache**: get a value from a tiingo query result set. The position of the cell invoking this function is used
to figure out which cell to get from the result set.

Parameters

* ``QueryKey``: should match the QueryKey given to `s2tiingo`.
* ``XOffset``: defaults to 0. If the left hand side of the result grid on your sheet is not column A this should
  be the number of columns across.
* ``YOffset``: defaults to 0. If the top row of the result grid on your sheet is not row 1 this should
  be the number of rows down.
* ``Trigger``: an optional trigger. The value isn't used inside the function, but a change in the input can
  be used to force repeat execution. 

Return value: a value from the result set, or #N/A.

**s2vtcache**: a volatile version of s2tcache.

Parameters

* ``QueryKey``: should match the QueryKey given to `s2tiingo`.
* ``XOffset``: defaults to 0. If the left hand side of the result grid on your sheet is not column A this should
  be the number of columns across.
* ``YOffset``: defaults to 0. If the top row of the result grid on your sheet is not row 1 this should
  be the number of rows down.

Return value: a value from the result set, or #N/A.

**s2baremetrics**: launch a Baremetrics metric query.

Parameters

* ``QueryKey``: a value or cell reference evaluating to a string that matches a value in column C of
  the s2cfg sheet. The s2cfg row with the matching column C value will be used to specify a Baremetrics query.
  See the baremetrics_summary1 or baremetrics_metric1 example sheets.
* ``Trigger``: an optional trigger. The value isn't used inside the function, but a change in the input can
  be used to force repeat execution. 
  
Return value: "OK" if the function succeeds, an Excel error otherwise.

**s2bcache**: get a value from a Baremetrics query result set. 

Parameters

* ``QueryKey``: should match the QueryKey given to `s2baremetrics`.
* ``Date``: Baremetrics result sets are keyed on date; think of date as picking out a row. You should supply a
  string in yyyy-MM-dd format, or use the `s2today` function. Don't use Excel's volatile `TODAY` function as 
  you'll cause an endless recalc cycle.
* ``Field``: pick out a column in the result set row selected by `Date`.
* ``Trigger``: an optional trigger. The value isn't used inside the function, but a change in the input can
  be used to force repeat execution. 

Return value: a value from the result set, or #N/A.

**s2vbcache**: a volatile version of s2bcache.

Parameters

* ``QueryKey``: should match the QueryKey given to `s2tiingo`.
* ``Date``: Baremetrics result sets are keyed on date; think of date as picking out a row. You should supply a
  string in yyyy-MM-dd format, or use the `s2today` function. Don't use Excel's volatile `TODAY` function as 
  you'll cause an endless recalc cycle.
* ``Field``: pick out a column in the result set row selected by `Date`.

Return value: a value from the result set, or #N/A.

**s2sub**: subscribe to RTD updates generated by s2cron, s2quandl or s2websock. 

Parameters

* ``SubCache``: [quandl|cron|websock]
* ``CacheKey``: should match the CronKey or QueryKey given to s2cron or s2quandl.
* ``Property``: [status|count|next|last|mX_Y_Z] count: cron event count for s2cron, rows in result set for s2quandl.
  next: time of next cron event. last: time of last cron event.

Return value: RTD value, or #N/A.

**s2websock**: subscribe via WebSockets to a page in a SpreadServe hosted sheet.

Parameters

* ``SockKey``: a value or cell reference evaluating to a string that matches a value in column C of
  the s2cfg sheet. The s2cfg row with the matching column C value will be used to specify the URL of
  a page in a SpreadServe hosted spreadsheet. See the websock1 example sheet.

Return value: "OK" if the function succeeds, an Excel error otherwise.

**s2twebsock**: subscribe via WebSockets to a Tiingo market data feed.

Parameters

* ``SockKey``: a value or cell reference evaluating to a string that matches a value in column C of
  the s2cfg sheet. The s2cfg row with the matching column C value will be used to specify the URL 
  for the Tiingo websocket connection. See the tiingows1 example sheet.

Return value: "OK" if the function succeeds, an Excel error otherwise.

**s2wscache**: get a value from a WebSocket subscription cache. 

Parameters

* ``SockKey``: should match the SockKey given to `s2websocket`.
* ``CellKey``: for instance, m2_6_0 for col 3, row 7 on first sheet. Use 'Page Source' in your browser to 
  examine the HTML on a page you want to subscribe to, and look for the div id tags to figure out the
  value you need.
* ``Trigger``: an optional trigger. 

Return value: a value from the cache, or #N/A.

**s2vwscache**: a volatile version of ``s2wscache``.

Parameters

* ``SockKey``: should match the SockKey given to `s2websocket`.
* ``CellKey``: for instance, m2_6_0 for col 3, row 7 on first sheet. Use 'Page Source' in your browser to 
  examine the HTML on a page you want to subscribe to, and look for the div id tags to figure out the
  value you need.

Return value: a value from the cache, or #N/A.

**s2today**: non volatile alternative to Excel's `TODAY`.

Parameters

* ``Offset``: 0 to get today, -1 for yesterday, +1 for tomorrow, -7 for a week ago, +7 for a week from now.

Return value: a yyyy-MM-dd formatted date string.

