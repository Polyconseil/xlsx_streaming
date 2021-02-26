import datetime
import logging
import re
from xml.etree import ElementTree as ETree


logger = logging.getLogger(__name__)


OPENXML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
OPENXML_NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
OPENXML_COLUMN_RE = re.compile(r'([A-Z]+)[0-9]+')


def _timezone_helper():
    _timezone_helper.timezone = None
    def get():
        return _timezone_helper.timezone
    def set_(value):
        _timezone_helper.timezone = value
    return get, set_


get_export_timezone, set_export_timezone = _timezone_helper()


def render_worksheet(rows_batches, openxml_sheet_string, encoding='utf-8'):
    """
        Render a collection of row batches to open xml.

        args:
            rows_batches (iterable): each element is a list of lists containing the row values
            openxml_sheet_string (str): a template for the final sheet containing the header and an example row
    """
    header = (
        '<worksheet xmlns="{ns}" xmlns:r="{ns_r}">\n'
        ' <sheetData>\n'
        .format(ns=OPENXML_NS, ns_r=OPENXML_NS_R)
    ).encode(encoding)
    footer = (
        ' </sheetData>\n'
        '</worksheet>\n'
    ).encode(encoding)

    yield header
    current_line = 1
    header_tree, row_template = get_header_and_row_template(openxml_sheet_string)  # pylint: disable=unbalanced-tuple-unpacking
    if header_tree is not None:
        yield ETree.tostring(header_tree, encoding=encoding)
        current_line += 1
    for rows in rows_batches:
        yield render_rows(rows, row_template, start_line=current_line, encoding=encoding)
        current_line += len(rows)
    yield footer


def render_rows(rows, row_template, start_line, encoding='utf-8'):
    """
        Return a collection of open xml rows as bytes.

        The output is a byte string like:

            <row r='2'>…</row>
            <row r='3'>…</row>
            …

        args:
            rows (list): a list of list containing the row values
            row_template (xml.ElementTree): a template used for each row
            start_line (int): the line of the first row in the returned xml

        ..note: This function updates row_template each time it is called.
    """
    return b'\n'.join(render_row(row, row_template, i, encoding) for i, row in enumerate(rows, start_line))


def render_row(row_values, row_template, line, encoding='utf-8'):
    """
        Return an openxml row as bytes using row_template as a model, and row_values for the values.

        args:
            row_values (list): the list of values to update the row
            row_template (xml.ElementTree): a template for the current row
            line (int): the line of the current row

        ..note: This function updates row_template each time it is called.
    """
    reset_memory = line < 3  # with a header, the first call to render_row can be with line = 2
    if row_template is None:
        row_template = get_default_template(row_values, reset_memory)
    cells = list(row_template)
    if len(cells) != len(row_values):
        logger.debug(
            '``len(row_values)`` do not match the number of cells in ``row_template``. '
            'Ignoring template (all cells will be stored as text).'
        )
        row_template = get_default_template(row_values, reset_memory)
        cells = list(row_template)

    for value, cell_template in zip(row_values, cells):
        update_cell(cell_template, line, value)
    row_template.set('r', str(line))
    return ETree.tostring(row_template, encoding=encoding)


def update_cell(cell, line, value):
    """
        Update cell with a new line and a new value.

        The value must be compatible with the cell type (numeric, boolean or text).
        The function raises if this is not the case.
        Updating a cell with a None value sets cell.text to the empty string.
    """
    column = get_column(cell)
    cell_type = cell.attrib.get('t', 'n')
    update_function = {
        'n': _update_numeric_cell,
        'b': _update_boolean_cell,
    }.get(cell_type, _update_text_cell)

    try:
        update_function(cell, value)
    except Exception as e:  # pylint: disable=broad-except
        args = e.args or ['']
        msg = "(column '%s', line '%s') data does not match template: %s" % (column, line, args[0])
        logger.debug(msg)
    cell.set('r', '%s%s' % (column, line))


def _update_boolean_cell(cell, value):
    if value is not None and not isinstance(value, bool):
        raise AttributeError("expected a boolean got %s." % value)
    next(child for child in cell if child.tag == 'v').text = '' if value is None else str(int(value))


def _update_numeric_cell(cell, value):
    if value is None:
        cell_text = ''
    else:
        try:
            cell_text = str(datetime_to_excel_datetime(value))
        except TypeError:
            cell_text = str(value)
        try:
            float(cell_text)
        except Exception as e:  # pylint: disable=broad-except
            raise AttributeError("expected a numeric or date like value got %s." % cell_text) from e
    next(child for child in cell if child.tag == 'v').text = cell_text


def _update_text_cell(cell, value):
    if cell.get('t') != 'inlineStr':
        # write all the strings 'inline' to avoid messing up with
        # a string reference file in the final xlsx file
        cell.clear()
        cell.set('t', 'inlineStr')
        ETree.SubElement(ETree.SubElement(cell, 'is'), 't')
    updated_text = '' if value is None else str(value)
    next(child for child in cell.iter() if child.tag == 't').text = updated_text


def datetime_to_excel_datetime(dt_obj, date_1904=False):
    """
        Convert a datetime object to an Excel serial date and time. The integer
        part of the number stores the number of days since the epoch and the
        fractional part stores the percentage of the day.
        Credits: Taken from https://github.com/jmcnamara/XlsxWriter
    """

    if date_1904:
        # Excel for Mac date epoch.
        epoch = datetime.datetime(1904, 1, 1)
    else:
        # Default Excel epoch.
        epoch = datetime.datetime(1899, 12, 31)

    # convert to the export timezone before making the datetimes naive (excel has no notion of timezone)
    if getattr(dt_obj, 'tzinfo', None) is not None:
        timezone = get_export_timezone()
        if timezone is not None:
            dt_obj = dt_obj.astimezone(timezone)
        dt_obj = dt_obj.replace(tzinfo=None)

    # We handle datetime .datetime, .date and .time objects but convert
    # them to datetime.datetime objects and process them in the same way.
    if isinstance(dt_obj, datetime.datetime):
        delta = dt_obj - epoch
    elif isinstance(dt_obj, datetime.date):
        dt_obj = datetime.datetime.fromordinal(dt_obj.toordinal())
        delta = dt_obj - epoch
    elif isinstance(dt_obj, datetime.time):
        dt_obj = datetime.datetime.combine(epoch, dt_obj)
        delta = dt_obj - epoch
    elif isinstance(dt_obj, datetime.timedelta):
        delta = dt_obj
    else:
        raise TypeError("Unknown or unsupported datetime type")

    # Convert a Python datetime.datetime value to an Excel date number.
    excel_time = (delta.days
                  + (float(delta.seconds)
                     + float(delta.microseconds) / 1E6)
                  / (60 * 60 * 24))

    # Special case for datetime where time only has been specified and
    # the default date of 1900-01-01 is used.
    if (not isinstance(dt_obj, datetime.timedelta)
            and dt_obj.isocalendar() == (1900, 1, 1)):
        excel_time -= 1

    # Account for Excel erroneously treating 1900 as a leap year.
    if not date_1904 and excel_time > 59:
        excel_time += 1

    return excel_time


def rm_namespace(xml_element):
    """Remove namespace in the ``xml_element`` and all of its children."""
    for el in xml_element.iter():
        if '}' in el.tag:
            el.tag = el.tag.split('}', 1)[1]  # strip all namespaces


def count_cells(row):
    """Count the number of cells in a openxml row"""
    return len(list(child for child in row if child.tag == 'c'))


def get_column(cell):
    """Return the column attribute ([A-Z]+) of the openxml cell."""
    match = re.match(OPENXML_COLUMN_RE, cell.get('r', ''))
    if match is None:
        raise AttributeError('The cell attribute is not a valid OpenXML column name: %s' % cell.get('r'))
    return match.group(1)


def _get_column_letter(col_idx):
    """Convert a column number into a column letter (3 -> 'C')

    Right shift the column col_idx by 26 to find column letters in reverse
    order.  These numbers are 1-based, and can be converted to ASCII
    ordinals by adding 64.

    Credits: openpyxl library
    """
    # these indicies corrospond to A -> ZZZ and include all allowed
    # columns
    if not 1 <= col_idx <= 18278:
        raise ValueError("Invalid column index {}".format(col_idx))
    letters = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx, 26)
        # check for exact division and borrow if needed
        if remainder == 0:
            remainder = 26
            col_idx -= 1
        letters.append(chr(remainder+64))
    return ''.join(reversed(letters))


def get_header_and_row_template(openxml_sheet):
    """
        Extract the header (potentially None) and the first row from openxml_sheet
        args:
            openxml_sheet (str): the openxml sheet imported as unicode
        return (tuple): a tuple of (header_tree, row_template_tree) ElementTree objects
    """
    tree = ETree.fromstring(openxml_sheet)
    rm_namespace(tree)
    sheet_data = next(child for child in tree if child.tag == 'sheetData')
    header_and_row_template = []
    for child in sheet_data:
        if child.tag == 'row':
            header_and_row_template.append(child)
        if len(header_and_row_template) == 2:
            header, row = header_and_row_template  # pylint: disable=unbalanced-tuple-unpacking
            if count_cells(header) != count_cells(row):
                logger.debug(
                    'Header and row template do not have the same number of cells. '
                    'Ignoring template (all cells will be stored as text).'
                )
                header_and_row_template = [None, None]
            break
    else:
        header_and_row_template = [None] + (header_and_row_template or [None])
    return tuple(header_and_row_template)


def get_default_template(row_values, reset_memory=False):
    """
        Return the default template row.
        It has only text cells, and as many columns as in row_values.
    """
    get_default_template._memoized = [] if reset_memory else getattr(get_default_template, '_memoized', [])
    if get_default_template._memoized:
        return get_default_template._memoized[0]

    root = ETree.Element('row', {'r': '1'})
    for i in range(1, len(row_values) + 1):
        el_t = ETree.Element('t')
        el_t.text = 'Default'
        el_is = ETree.Element('is')
        el_is.append(el_t)
        cell = ETree.Element('c', {'r': '%s%s' % (_get_column_letter(i), 1), 't': 'inlineStr'})
        cell.append(el_is)
        root.append(cell)

    get_default_template._memoized.append(root)
    return root
