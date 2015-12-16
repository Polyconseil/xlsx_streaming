# -*- coding: utf-8 -*-

from __future__ import unicode_literals

import zipfile

import zipstream

from . import render

TEMPLATE_SHEET_NAME = 'xl/worksheets/sheet1.xml'


def serialize_queryset_by_batch(qs, serializer, batch_size):
    start, last_batch = 0, []
    while start == 0 or len(last_batch) == batch_size:
        last_batch = list(qs[start:start + batch_size])  # force queryset evaluation
        yield serializer(last_batch)
        start += batch_size


def zip_to_zipstream(zip_file, only=None, exclude=None):
    """
    args:
        zip_file (ZipFile): the original zipfile.ZipFile
        only (list): the file names of the files to be included in the stream
        exclude (list): the file names of the files to be excluded in the stream
        ..note: only and exclude cannot be used at the same time
    """
    only = only or []
    exclude = exclude or []
    if only and exclude:
        raise AttributeError('`only` and `exclude` cannot be used at the same time')

    file_names = zip_file.namelist()
    if only:
        file_names = [name for name in file_names if name in only]
    elif exclude:
        file_names = [name for name in file_names if name not in exclude]

    zip_stream = zipstream.ZipFile(mode='w', compression=zipstream.ZIP_DEFLATED)
    for file_name in file_names:
        zip_stream.write_iter(
            arcname=file_name,
            iterable=iter([zip_file.read(file_name)]),
            compress_type=zipstream.ZIP_DEFLATED,
        )
    return zip_stream


def stream_queryset_as_xlsx(
        qs,
        xlsx_template,
        serializer=None,
        batch_size=1000,
        encoding='utf-8',
    ):
    """
        Iterate over qs by batch (typically a Django queryset) and stream the bytes of the
        xlsx document generated from the data.

        This function can typically be used to stream data extracted from a database by batch,
        making many small database requests instead of one big request which could timeout.

        args:
            qs (iterable): an iterable containing the rows (typically a Django queryset)
            xlsx_template (file like): an in memory xlsx file template containing
                the header (optional) and the first row
            serializer (function): a function applied to each batch of rows to transform them before saving
                them to the xlsx document (defaults to identity)
            batch_size (int): the size of each batch of rows

    ..note: If the xlsx template contains more than one row, the first row is kept as is in the final
        xlsx file (header row), and the second one is used as a template for all the generated rows.
    """
    sheet_name = TEMPLATE_SHEET_NAME
    serializer = serializer or (lambda x: x)

    batches = serialize_queryset_by_batch(qs, serializer=serializer, batch_size=batch_size)

    zip_template = zipfile.ZipFile(xlsx_template, mode='r')
    xlsx_sheet_string = zip_template.read(sheet_name).decode(encoding)

    zipped_stream = zip_to_zipstream(zip_template, exclude=[sheet_name])
    # Write the generated worksheet to the stream
    worksheet_stream = render.render_worksheet(batches, xlsx_sheet_string, encoding)
    zipped_stream.write_iter(
        arcname=sheet_name,
        iterable=worksheet_stream,
        compress_type=zipstream.ZIP_DEFLATED
    )

    return zipped_stream
