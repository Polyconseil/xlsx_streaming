xlsx_streaming
================

.. image:: https://badge.fury.io/py/xlsx_streaming.svg
    :target: http://badge.fury.io/py/xlsx_streaming
.. image:: https://travis-ci.org/Polyconseil/xlsx_streaming.svg?branch=master
    :target: https://travis-ci.org/Polyconseil/xlsx_streaming

``xlsx_streaming`` library lets you stream data from your database to an xlsx document. The xlsx document can hence be streamed over HTTP without the need to store it in memory on the server. Database queries are made by batch, making it safe to export even very large tables.

Bug reports, patches and suggestions welcome! Just open an issue_ or send a `pull request`_.

tests
-----

After having installed the development requirements, just run::

    python -m unittest discover

.. _issue: https://github.com/sebdiem/xlsx_streaming/issues/new
.. _pull request: https://github.com/sebdiem/xlsx_streaming/compare/
