First steps
===========

The ``xlsx_streaming`` library is primarily meant to be used with streaming HTTP response objects.
It provides an API to stream data extracted from a database directly into an xlsx document.

This library does not depend on any web framework or any python excel library.
However it was primarily designed for django and has not been tested with any other web framework yet.

Django example
++++++++++++++

.. code:: python 

    from django.http import StreamingHttpResponse

    import xlsx_streaming

    def my_view(request):
        qs = MyModel.objects.all()
        with open('template.xlsx') as template:
            stream = xlsx_streaming.stream_queryset_as_xlsx(
                qs,
                template,
                batch_size=50
            )

        response = StreamingHttpResponse(
            stream,
            content_type='application/vnd.xlsxformats-officedocument.spreadsheetml.sheet',
        )
        response['Content-Disposition'] = 'attachment; filename=test.xlsx'
        return response

With this kind of code, your database will not be hit too hard by the file export, even when the export's size is several dozens of megabytes.
Like in the example above, you can use a static template, but you can also generate the template dynamically using the python excel library of your choice (`openpyxl`_, …).

.. _openpyxl: https://openpyxl.readthedocs.org/en/default/

Built-in serialization
======================

The ``xlsx_streaming`` library provides builtin support for numerical, date/datetime/timedelta, and boolean conversion to excel equivalent.
All other data types are converted to text by default.

Using custom serializers
========================

If you want to preprocess the data from the database before inserting it in the xlsx document, you can pass a serializer argument to ``stream_queryset_as_xlsx``.
This function is applied to each queryset slice before writing the data to excel.
For example, with a Django Rest Framework serializer:

.. code:: python

    class MySerializer(serializers.Serializer):
        …

    serializer = lambda x: [d.values() for d in MySerializer(x, many=True).data]
    xlsx_streaming.stream_queryset_as_xlsx(qs, template, serializer=serializer)

Specifying the timezone of the export
=====================================

The datetime objects might be stored as 'UTC' inside your database, but you may want to see datetime objects exported in a specific timezone in the excel document.
Excel does not offer support for timezone. However, the ``xlsx_streaming`` library lets you specify the timezone to be used for exports (by default the timezone is kept unchanged).
Before writing a datetime, it will first be localized in the specified timezone before being converted to a naive datetime.
This functionality requires the `pytz`_ library.

For example, if you set:

.. code:: python

    from pytz import timezone

    xlsx_streaming.set_export_timezone(timezone('US/Eastern'))

.. _pytz: http://pytz.sourceforge.net/

all the datetimes written with the ``xlsx_streaming`` library will first be localized in the 'US/Eastern' timezone.
