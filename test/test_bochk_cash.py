"""
Test the open_jpm.py
"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from bochk.utility import get_current_path
from bochk.open_bochk import read_cash_fields, read_cash_line, read_cash_bochk, \
                                read_holdings_bochk, get_cash_date_as_string



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



    def test_read_cash_bochk1(self):
        filename = get_current_path() + '\\samples\\sample_cash1.xls'
        port_values = {}
        read_cash_bochk(filename, port_values)
        cash_entries = port_values['cash']
        cash_transactions = port_values['cash_transactions']
        self.verify_cash1(cash_entries, cash_transactions)



    def test_read_cash_bochk2(self):
        filename = get_current_path() + '\\samples\\sample_cash2.xls'
        port_values = {}
        read_cash_bochk(filename, port_values)
        cash_entries = port_values['cash']
        cash_transactions = port_values['cash_transactions']
        self.verify_cash2(cash_entries, cash_transactions)



    def test_get_cash_date_as_string(self):
        cash_file = get_current_path() + '\\samples\\sample_cash2.xls'
        holdings_file = get_current_path() + '\\samples\\sample_holdings2.xls'
        port_values = {}
        read_cash_bochk(cash_file, port_values)
        read_holdings_bochk(holdings_file, port_values)
        cash_entry = port_values['cash'][0]
        d = get_cash_date_as_string(port_values, cash_entry)
        self.assertEqual(d, '2016-11-16')



    def test_get_cash_date_as_string2(self):
        cash_file = get_current_path() + '\\samples\\sample_holdings5_cash.xls'
        holdings_file = get_current_path() + '\\samples\\sample_holdings5.xls'
        port_values = {}
        read_cash_bochk(cash_file, port_values)
        read_holdings_bochk(holdings_file, port_values)
        cash_entry = port_values['cash'][0]
        d = get_cash_date_as_string(port_values, cash_entry)
        self.assertEqual(d, '2016-7-6')



    def test_get_cash_date_as_string3(self):
        cash_file = get_current_path() + '\\samples\\sample_holdings4_cash.xls'
        holdings_file = get_current_path() + '\\samples\\sample_holdings4.xls'
        port_values = {}
        read_cash_bochk(cash_file, port_values)
        read_holdings_bochk(holdings_file, port_values)
        cash_entry = port_values['cash'][0]
        d = get_cash_date_as_string(port_values, cash_entry)
        self.assertEqual(d, '2016-7-6')



    def verify_cash1(self, cash_entries, cash_transactions):
        """
        For sample_cash1.xls
        """
        self.assertEqual(len(cash_entries), 3)
        for entry in cash_entries:
            if entry['Currency'] == 'HKD':
                self.verify_cash_entry01(entry)
            elif entry['Currency'] == 'USD':
                self.verify_cash_entry02(entry)

        self.assertEqual(len(cash_transactions), 3)
        self.verify_cash_transaction02(cash_transactions[2])



    def verify_cash2(self, cash_entries, cash_transactions):
        """
        For sample_cash2.xls
        """
        self.assertEqual(len(cash_entries), 3)
        for entry in cash_entries:
            if entry['Currency'] == 'HKD':
                self.verify_cash_entry1(entry)
            elif entry['Currency'] == 'USD':
                self.verify_cash_entry2(entry)

        self.assertEqual(len(cash_transactions), 0)



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
        The first entry in sample_cash1.xls
        """
        self.assertEqual(len(cash_entry), 8)
        self.assertEqual(cash_entry['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_entry['Currency'], 'HKD')
        self.assertEqual(cash_entry['Current Ledger Balance'], 0)
        self.assertEqual(cash_entry['Current Available Balance'], 0)



    def verify_cash_entry02(self, cash_entry):
        """
        The last (5th) entry in sample_cash1.xls
        """
        self.assertEqual(len(cash_entry), 9)
        self.assertEqual(cash_entry['Account Number'], '\'01287508062518')
        self.assertEqual(cash_entry['Account Type'], 'USD Current Account')
        self.assertEqual(cash_entry['Currency'], 'USD')
        self.assertEqual(cash_entry['Ledger Balance'], 62732794.47)
        self.assertEqual(cash_entry['Current Available Balance'], 37598962.52)



    def verify_cash_transaction02(self, cash_transaction):
        """
        from the last (5th) entry in sample_cash1.xls
        """
        self.assertEqual(len(cash_transaction), 12)
        self.assertEqual(cash_transaction['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_transaction['Account Number'], '\'01287508062518')
        self.assertEqual(cash_transaction['Currency'], 'USD')
        self.assertEqual(cash_transaction['Processing Date / Time'], datetime(2016,7,11))
        self.assertAlmostEqual(cash_transaction['Amount'], 3176345.83)
        self.assertEqual(cash_transaction['Transaction Reference'], '49969')