"""
Test the open_jpm.py
"""

import unittest2
import datetime
from xlrd import open_workbook
from bochk.utility import get_current_path
from bochk.open_bochk import field_begins, read_fields



class TestBOCHK(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestBOCHK, self).__init__(*args, **kwargs)

    def setUp(self):
        """
            Run before a test function
        """
        pass



    def tearDown(self):
        """
            Run after a test finishes
        """
        pass



    def test_field_begins(self):
        filename = get_current_path() + '\\samples\\sample_holdings2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_index(0)
        row = 0
        self.assertFalse(field_begins(ws, row))
        row = 2
        self.assertTrue(field_begins(ws, row))


    def test_read_fields(self):
        """
        Read the date
        """
        filename = get_current_path() + '\\samples\\sample_holdings2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_index(0)
        row = 0

        while not field_begins(ws, row):
            row = row + 1

        fields = read_fields(ws, row)
        self.assertEqual(len(fields), 21)
        self.assertEqual(fields[0], 'record_type')
        self.assertEqual(fields[11], 'holding_quantity')
        self.assertEqual(fields[20], 'equivalent_market_value')
