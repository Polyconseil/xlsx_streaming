0.4.1 (2019-11-07)
------------------

- Fix bug in ``stream_queryset_as_xlsx()`` that was introduced in
  version 0.4.0. When called with a Django QuerySet object (or
  something similar), the function made a single SQL query, instead of
  multiple SQL queries depending on the ``batch_size`` argument. This
  bug may cause performance issues when fetching many rows from the
  database.


0.4.0 (2019-11-06)
------------------

- ``stream_queryset_as_xlsx()`` now accepts a generator. [Hugo Lecuyer].


0.3.1 (2017-02-22)
------------------

* First sheet of a workbook is now more reliably detected

0.3.0 (2017-02-22)
------------------

* It is now possible to stream data to an Excel document without providing
  a template. In this case all cells are formatted as text (no date or number
  formatting). If there is an error with the template provided by the user,
  fall back to this behaviour.
* Improve error handling (try to recover and log instead of raise)
* Change license

0.2.0 (2015-12-16)
------------------

* Add python 2.6 compatibility

0.1.0 (2015-12-16)
------------------

* First public version.
