# -*- coding: utf-8 -*-

from __future__ import unicode_literals

import datetime
from io import BytesIO
import zipfile

import openpyxl

from xlsx_streaming import streaming

def gen_xlsx_template(with_header=False):
    wb = openpyxl.Workbook()
    rows = [[42, 'éOui>€', datetime.datetime(2012, 1, 2, 10, 10)]]
    if with_header:
        rows = [['Id', 'Description', 'Date']] + rows
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            wb.active.cell(row=i + 1, column=j + 1).value = value
    return BytesIO(openpyxl.writer.excel.save_virtual_workbook(wb))

def gen_xlsx_sheet(with_header=False):
    xlsx_template = gen_xlsx_template(with_header=with_header)
    zip_template = zipfile.ZipFile(xlsx_template, mode='r')
    sheet_name = streaming.get_first_sheet_name(zip_template)
    return zip_template.read(sheet_name).decode('utf-8')
