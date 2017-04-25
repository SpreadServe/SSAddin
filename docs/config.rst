SpreadServe Addin Configuration
===============================

**Log files**

The SSAddin creates log files in your ``%TEMP%`` directory. To find them do this in a DOS box::

    echo %TEMP%
    cd %TEMP%
    dir ssaddin* /od
    dir *.csv
    
Note that the process ID of the Excel instance hosting your SSAddin is embedded in the log file
name. The log file captures all the RTD updates sent by the addin to the sheet, together with
their values. It also logs the start and end of Quandl queries. The addin also dumps the result
sets returned from Quandl into CSV files in the ``%TEMP%`` directory. The files are named
``<QueryKey>_<ProcessID>.csv``.

**s2cfg sheet**

Any spreadsheet that uses SSAddin must have a sheet called ``s2cfg``. The SSAddin worksheet
functions get their configuration from the s2cfg sheet and will fail if it doesn't exist
or if its contents are not correctly laid out. The log files should alert you if there's a
problem in your s2cfg sheet. They are also a good way of checking that the addin has composed
your quandl or tiingo queries as you expected. Bear in mind these points on how the addin
scans the s2cfg sheet for configuration. Also check the example sheets in the xls sub directory
for concrete illustrations of the guidelines below.

* SSAddin scans the s2cfg sheet from the first row downwards. It will stop scanning when it
  finds a row with an empty cell in column A. This means you can't have spaces between your
  config. It must all be in a single contiguous block from row 1 downwards.
* The value in column A must be ``quandl``, ``tiingo``, ``cron``, ``websock`` or ``twebsock``.
* Depending on the value in column A there are different expectations for the values in
  column B onwards.
  
  * ``quandl``: column B should be ``query`` or ``config``
  
    * ``query``: column C should be the unique QueryKey that's passed to the ``s2quandl``
      function, column D should be ``dataset`` and column E should name a Quandl dataset
      eg ``FRED/DED1`` or ``OPEC/ORB``. Any further columns should give key value pairs
      to tacked on to the Quandl query URL after the ``?``  For instance column F could be
      ``rows`` and column G ``5`` so that ``?rows=5`` is appended to the URL query submitted
      to quandl.
    * ``config``: column pairs from  C & D onwards are reserved for name value pairs that
      apply to all queries. Currently only ``auth_token`` is supported. If you put ``auth_token``
      in column C, then put your actual key in column D for it to be added to all queries.
      However, we recommend you put your key in SSAddin.xll.config instead, so you don't 
      indavertently share your key when sharing your spreadsheet.
  
  * ``tiingo``: column B should be ``query`` or ``config``
  
    * ``query``: column C should be the unique QueryKey that's passed to the ``s2tiingo``
      function, column D should be ``ticker`` and column E should be a ticker symbol
      eg ``msft`` or ``aapl``. The ticker symbol should be lower case. Column F should
      be ``root``, followed by ``daily`` or ``funds`` in column G. Column H is optional.
      If it's present it should be ``leaf`` and then column I should be ``prices``. If
      it's absent a tiingo query that gets meta data for the symbol will be dispatched.
      Finally, columns J, K, L & M can be used to specify startDate and endDate for
      historical price queries. 
    * ``config``: column pairs from  C & D onwards are reserved for name value pairs that
      apply to all queries or Tiingo web socket connections (see twebsock below). 
      Supported config keys are...
      
      * ``auth_token``: put ``auth_token`` in column C, and your actual key in column D
        for it to be added to all queries or used by twebsock.
      * ``http_proxy_host``: if this appears in column C then column D should give a proxy
        hostname. SSAddin will then connect via the proxy rather than direct to the internet.
      * ``http_proxy_port``: port for the proxy connection.
      * ``http_proxy_user``: user name for the proxy connection. Often this is in DOMAIN\USER
        format for Windows Active Directory user IDs.
      * ``http_proxy_password``: password for the proxy connection.
      
  * ``baremetrics``: column B should be ``query`` or ``config``
  
    * ``query``: column C should be the unique QueryKey that's passed to the ``s2baremetrics``
      function, column D should be ``qtype`` and column E should be a ``summary``, ``plan``
      or ``metric``. For a qtype of ``plan`` or ``metric`` you need a following key/value
      pair that specifies which metric. The key, in column F should be ``metric``, and
      then in column G you should specify ``mrr``, `arpu``, ``ltv`` or any of the available
      metrics. Finally, columns J, K, L & M can be used to specify ``start_date`` and ``end_date``
      with the date values in columns K and M supplied by ``s2today`` or handcoded yyy-MM-dd
      strings. Don't use Excel's own `TODAY` function for these as it's volatile and will
      cause an endless calc cycle.
    * ``config``: column pairs from  C & D onwards are reserved for name value pairs that
      apply to all queries or Tiingo web socket connections (see twebsock below). 
      Supported config keys are...
      
      * ``auth_token``: put ``auth_token`` in column C, and your actual key in column D
        for it to be added to all queries or used by twebsock.
      * ``http_proxy_host``: if this appears in column C then column D should give a proxy
        hostname. SSAddin will then connect via the proxy rather than direct to the internet.
      * ``http_proxy_port``: port for the proxy connection.
      * ``http_proxy_user``: user name for the proxy connection. Often this is in DOMAIN\USER
        format for Windows Active Directory user IDs.
      * ``http_proxy_password``: password for the proxy connection.      
      
  * ``twebsock``: when column B contains ``tiingo`` then column C specifies a ``SockKey`` to pass
    to ``s2twebsock``. Column D should give the URL for the Tiingo API socket eg ``wss://api.tiingo.com/iex``
  
  * ``cron``: when column B contains ``tab`` then column C should have a unique ``CronKey``
    that will be passed to the ``s2cron`` worksheet function which will then get the cron
    job specification from columns D to K. This job spec is then passed to SSAddin's internal
    `NCrontab <https://code.google.com/p/ncrontab/wiki/CrontabExamples>`_ implementation.
    Bear in mind that SSAddin uses a hacked version of NCrontab that extends the spec to
    add seconds.
    
    * ``D``: seconds
    * ``E``: minutes
    * ``F``: hours
    * ``G``: days
    * ``H``: month
    * ``I``: weekday
    * ``J``: start - defaults to the start of today, today being the day when the process started.
    * ``K``: end - defaults to the end of today
    
  * ``websock``: when column B contains ``url`` then column C specifies a ``SockKey`` to pass
    to ``s2websock``. Column D should give the hostname of a SpreadServe server, column E the
    port number, and column F the rest of the URL, often referred to as the path.
    
Note that if column B has any other value than described above it will be ignored. One convention
you'll see in the SSAddin example s2cfg sheets is ``comment`` occurring in column B so that the
rest of the row can be used as headers to describe the real values below.
