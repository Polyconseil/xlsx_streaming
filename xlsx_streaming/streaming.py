import collections.abc
from itertools import chain, islice
import logging
import zipfile

import zipstream

from . import render
from .xlsx_template import DEFAULT_TEMPLATE


logger = logging.getLogger(__name__)

EXCEL_WORKSHEETS_PATH = 'xl/worksheets/'


def stream_queryset_as_xlsx(
        qs,
        xlsx_template=None,
        serializer=None,
        batch_size=1000,
        encoding='utf-8',
    ):
    """
    Iterate over qs by batch (typically a Django queryset) and stream the bytes of the
    xlsx document generated from the data.

    This function can typically be used to stream data extracted from a database by batch,
    making many small database requests instead of one big request which could timeout.

    Args:
        qs (Iterable): an iterable containing the rows (typically a Django queryset)
        xlsx_template (Optional[BytesIO]): an in memory xlsx file template containing
            the header (optional) and the first row used to infer data types for each column.
            If not provided, all cells will be formatted as text.
        serializer (Optional[Callable]): a function applied to each batch of rows to transform
            them before saving them to the xlsx document (defaults to identity).
        batch_size (Optional[int]): the size of each batch of rows
        encoding (Optional[str]): the file encoding

    Returns:
        Iterable: A streamable xlsx file

    Note:
        If the xlsx template contains more than one row, the first row is kept as is in the final
        xlsx file (header row), and the second one is used as a template for all the generated rows.
    """
    serializer = serializer or (lambda x: x)

    batches = serialize_queryset_by_batch(qs, serializer=serializer, batch_size=batch_size)

    try:
        zip_template = zipfile.ZipFile(xlsx_template, mode='r')
    except Exception:  # pylint: disable=broad-except
        logger.debug('Template is not a valid Excel file, ignoring it. Every cell will be saved as text.')
        zip_template = zipfile.ZipFile(DEFAULT_TEMPLATE, mode='r')

    sheet_name = get_first_sheet_name(zip_template)
    if sheet_name is None:
        logger.debug('Template is not a valid Excel file, ignoring it. Every cell will be saved as text.')
        zip_template = zipfile.ZipFile(DEFAULT_TEMPLATE, mode='r')
        sheet_name = get_first_sheet_name(zip_template)

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


def serialize_queryset_by_batch(qs, serializer, batch_size):
    if isinstance(qs, collections.abc.Iterator):
        qs_slices = _chunks(qs, batch_size)
        for batch in qs_slices:
            yield serializer(list(batch))
    else:
        start, last_batch = 0, []
        while start == 0 or len(last_batch) == batch_size:
            last_batch = list(qs[start:start + batch_size])  # force queryset evaluation
            yield serializer(last_batch)
            start += batch_size


def _chunks(iterable, size):
    iterator = iter(iterable)
    for first in iterator:
        yield chain([first], islice(iterator, size - 1))


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


def get_first_sheet_name(xlsx_zipfile):
    try:
        return next(
            path
            for path in xlsx_zipfile.namelist()
            if path.startswith(EXCEL_WORKSHEETS_PATH) and path.endswith('.xml')
        )
    except StopIteration:
        return None
