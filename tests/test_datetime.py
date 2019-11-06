import datetime
import unittest

from pytz import timezone

from xlsx_streaming import render


class TestDatetime(unittest.TestCase):
    def tearDown(self):
        render.set_export_timezone(timezone('UTC'))
        super().tearDown()

    def test_set_timezone(self):
        render.set_export_timezone(timezone('US/Eastern'))
        self.assertEqual(render.get_export_timezone(), timezone('US/Eastern'))

    def test_datetime_to_excel_no_timezone_change(self):
        excel_value = 40910.0  # 2012, 1, 2, 0, 0
        dt = timezone('UTC').localize(datetime.datetime(2012, 1, 2, 0, 0))
        self.assertEqual(render.datetime_to_excel_datetime(dt), excel_value)

    def test_datetime_to_excel_with_timezone_change(self):
        # test that the datetime is converted to the wanted timezone
        excel_value = 40910.0  # 2012, 1, 2, 0, 0
        dt = timezone('US/Eastern').localize(datetime.datetime(2012, 1, 2, 0, 0))
        dt = dt.astimezone(timezone('UTC'))

        render.set_export_timezone(timezone('US/Eastern'))
        # our datetime is now in UTC but we want the value in excel to be the US/Eastern equivalent
        self.assertEqual(render.datetime_to_excel_datetime(dt), excel_value)
