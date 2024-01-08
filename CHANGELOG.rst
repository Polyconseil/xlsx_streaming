1.2.0 (unreleased)
------------------

- Add support for Python 3.10 & 3.11 & 3.12
- Drop support for Python 3.6 & 3.7
- If provided template contains ``sheetViews`` with a ``pane`` element, this
  element will now be copied to the resulting Excel document.


1.1.0 (2021-06-23)
------------------

- ``render_worksheet()`` now allows rendering worksheet with iterators.


1.0.0 (2021-03-01)
------------------

- ``stream_queryset_as_xlsx()`` now accepts an iterator.
- |backward-incompatibility| Drop support of Python 2, 3.4 and 3.5. xlxs_streaming now supports
  only Python 3.6 and onward.


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


.. role:: raw-html(raw)

.. |backward-incompatibility| raw:: html

    <span style="background-color: #ffffbc; padding: 0.3em; font-weight: bold;">backward incompatibility</span>
