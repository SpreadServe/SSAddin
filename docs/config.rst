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
your quandl queries as you expected. Bear in mind these points on how the addin scans the s2cfg
sheet for configuration...

* SSAddin scans the s2cfg sheet from the first row downwards. It will stop scanning when it
  finds a row with an empty cell in column A. This means you can't have spaces between your
  config. It must all be in a single contiguous block from row 1 downwards.
* The value in column A must be ``quandl``, ``cron`` or ``websock``.
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
      apply to all queries. Currently only ``auth_key`` is supported. If you put ``auth_key``
      in column C, then put your actual key in column D for it to be added to all queries.
  
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
