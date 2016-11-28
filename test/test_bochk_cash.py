"""
Test the open_jpm.py
"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from bochk.utility import get_current_path
from bochk.open_bochk import read_cash_fields, read_cash_line



class TestBOCHKCash(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestBOCHKCash, self).__init__(*args, **kwargs)

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



    def get_worksheet(self, filename):
        filename = get_current_path() + filename
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_index(0)
        return ws



    def test_read_cash_fields(self):
        """
        Read the date
        """
        ws = self.get_worksheet('\\samples\\sample_cash2.xls')
        row = 0

        cell_value = ws.cell_value(row, 0)
        while row < ws.nrows:
            if isinstance(cell_value, str) and cell_value.strip() == 'Account Name':
                break
            row = row + 1

        fields = read_cash_fields(ws, row)
        self.assertEqual(len(fields), 18)
        self.assertEqual(fields[0], 'Account Name')
        self.assertEqual(fields[3], 'Currency')
        self.assertEqual(fields[14], 'Ledger Balance')
        self.assertEqual(fields[17], 'Cheque Number')



    def test_read_cash_line(self):
        """
        Read the date
        """
        ws = self.get_worksheet('\\samples\\sample_cash2.xls')
        row = 0
        port_values = {}
    
        cell_value = ws.cell_value(row, 0)
        while row < ws.nrows:
            if isinstance(cell_value, str) and cell_value.strip() == 'Account Name':
                break
            row = row + 1

        fields = read_cash_fields(ws, row)
        cash_entry, cash_transaction = read_cash_line(ws, row+1, fields)
        self.assertTrue(cash_transaction is None)
        self.verify_cash_entry1(cash_entry)

        cash_entry, cash_transaction = read_cash_line(ws, row+3, fields)
        self.assertTrue(cash_transaction is None)
        self.verify_cash_entry2(cash_entry)



    def test_read_cash_line2(self):
        """
        Read the date
        """
        ws = self.get_worksheet('\\samples\\sample_cash1.xls')
        row = 0
        port_values = {}
    
        cell_value = ws.cell_value(row, 0)
        while row < ws.nrows:
            if isinstance(cell_value, str) and cell_value.strip() == 'Account Name':
                break
            row = row + 1

        fields = read_cash_fields(ws, row)
        cash_entry, cash_transaction = read_cash_line(ws, row+1, fields)
        self.assertTrue(cash_transaction is None)
        self.verify_cash_entry01(cash_entry)

        cash_entry, cash_transaction = read_cash_line(ws, row+5, fields)
        self.verify_cash_entry02(cash_entry)
        self.verify_cash_transaction02(cash_transaction)



    def verify_cash_entry1(self, cash_entry):
        """
        The first entry in sample_cash2.xls
        """
        self.assertEqual(len(cash_entry), 8)
        self.assertEqual(cash_entry['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_entry['Currency'], 'HKD')
        self.assertEqual(cash_entry['Current Ledger Balance'], 0)
        self.assertEqual(cash_entry['Current Available Balance'], 0)



    def verify_cash_entry2(self, cash_entry):
        """
        The last (3rd) entry in sample_cash2.xls
        """
        self.assertEqual(len(cash_entry), 8)
        self.assertEqual(cash_entry['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_entry['Currency'], 'USD')
        self.assertAlmostEqual(cash_entry['Current Ledger Balance'], 94233724.32)
        self.assertAlmostEqual(cash_entry['Current Available Balance'], 94233724.32)



    def verify_cash_entry01(self, cash_entry):
        """
        The first entry in sample_cash2.xls
        """
        self.assertEqual(len(cash_entry), 8)
        self.assertEqual(cash_entry['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_entry['Currency'], 'HKD')
        self.assertEqual(cash_entry['Current Ledger Balance'], 0)
        self.assertEqual(cash_entry['Current Available Balance'], 0)



    def verify_cash_entry02(self, cash_entry):
        """
        The last (5th) entry in sample_cash2.xls
        """
        self.assertEqual(len(cash_entry), 9)
        self.assertEqual(cash_entry['Account Number'], '\'01287508062518')
        self.assertEqual(cash_entry['Account Type'], 'USD Current Account')
        self.assertEqual(cash_entry['Currency'], 'USD')
        self.assertEqual(cash_entry['Ledger Balance'], 62732794.47)
        self.assertEqual(cash_entry['Current Available Balance'], 37598962.52)



    def verify_cash_transaction02(self, cash_transaction):
        """
        from the last (5th) entry in sample_cash2.xls
        """
        self.assertEqual(len(cash_transaction), 12)
        self.assertEqual(cash_transaction['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_transaction['Account Number'], '\'01287508062518')
        self.assertEqual(cash_transaction['Currency'], 'USD')
        self.assertEqual(cash_transaction['Processing Date / Time'], datetime(2016,7,11))
        self.assertAlmostEqual(cash_transaction['Amount'], 3176345.83)
        self.assertEqual(cash_transaction['Transaction Reference'], '49969')