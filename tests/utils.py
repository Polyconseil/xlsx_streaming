import datetime
import io
import tempfile
import zipfile

import openpyxl

from xlsx_streaming import streaming


def gen_xlsx_template(with_header=False, with_views=False):
    wb = openpyxl.Workbook()
    rows = [[42, 'éOui>€', datetime.datetime(2012, 1, 2, 10, 10)]]
    if with_header:
        rows = [['Id', 'Description', 'Date']] + rows
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            wb.active.cell(row=i + 1, column=j + 1).value = value
    if with_views:
        wb.active.freeze_panes = "A2"
    with tempfile.NamedTemporaryFile() as fp:
        openpyxl.writer.excel.save_workbook(wb, fp)
        fp.seek(0)
        return io.BytesIO(fp.read())


def gen_xlsx_sheet(with_header=False, with_views=False):
    xlsx_template = gen_xlsx_template(with_header=with_header, with_views=with_views)
    with zipfile.ZipFile(xlsx_template, mode='r') as zip_template:
        sheet_name = streaming.get_first_sheet_name(zip_template)
        return zip_template.read(sheet_name).decode('utf-8')
