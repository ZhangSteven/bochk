"""
Test the open_jpm.py
"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from bochk.utility import get_current_path
from bochk.open_bochk import read_cash_fields, read_cash_line, read_cash_bochk, \
                                read_holdings_bochk, consolidate_cash



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
        ws = self.get_worksheet('\\samples\\sample_cash2 _ 16112016.xls')
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
        ws = self.get_worksheet('\\samples\\sample_cash2 _ 16112016.xls')
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
        ws = self.get_worksheet('\\samples\\sample_cash1 _ 16112016.xls')
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
        filename = get_current_path() + '\\samples\\sample_cash1 _ 16112016.xls'
        port_values = {}
        read_cash_bochk(filename, port_values)
        cash_entries = port_values['cash']
        cash_transactions = port_values['cash_transactions']
        self.verify_cash1(cash_entries, cash_transactions)



    def test_read_cash_bochk2(self):
        filename = get_current_path() + '\\samples\\sample_cash2 _ 16112016.xls'
        port_values = {}
        read_cash_bochk(filename, port_values)
        cash_entries = port_values['cash']
        cash_transactions = port_values['cash_transactions']
        self.verify_cash2(cash_entries, cash_transactions)



    def test_read_cash_bochk3(self):
        # the in house fund, it has cash consoliation (two HKD accounts, savings
        # and current account)
        filename = get_current_path() + '\\samples\\Cash _ 31082017.xlsx'
        port_values = {}
        read_cash_bochk(filename, port_values)
        consolidate_cash(port_values)
        self.verify_cash3(port_values['cash'], port_values['cash_transactions'])



    def test_read_cash_bochk4(self):
        # the in house fund, it has cash consoliation (two HKD accounts, savings
        # and current account)
        filename = get_current_path() + '\\samples\\Cash Stt _30042018.xlsx'
        port_values = {}
        read_cash_bochk(filename, port_values)
        consolidate_cash(port_values)
        self.verify_cash4(port_values['cash'], port_values['cash_transactions'])



    def verify_cash1(self, cash_entries, cash_transactions):
        """
        For sample_cash1 _ 16112016.xls
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
        For sample_cash2 _ 16112016.xls
        """
        self.assertEqual(len(cash_entries), 3)
        for entry in cash_entries:
            if entry['Currency'] == 'HKD':
                self.verify_cash_entry1(entry)
            elif entry['Currency'] == 'USD':
                self.verify_cash_entry2(entry)

        self.assertEqual(len(cash_transactions), 0)



    def verify_cash3(self, cash_entries, cash_transactions):
        """
        For samples/Cash _ 31082017.xlsx
        """
        self.assertEqual(len(cash_entries), 3)
        for entry in cash_entries:
            if entry['Currency'] == 'HKD':
                self.assertAlmostEqual(entry['Current Ledger Balance'], 253456.32)
                self.assertAlmostEqual(entry['Current Available Balance'], 253456.32)
                self.assertFalse('Ledger Balance' in entry)

            elif entry['Currency'] == 'USD':
                self.assertAlmostEqual(entry['Current Ledger Balance'], 284041.93)
                self.assertAlmostEqual(entry['Current Available Balance'], 284041.93)
                self.assertFalse('Ledger Balance' in entry)

        self.assertEqual(len(cash_transactions), 0)




    def verify_cash4(self, cash_entries, cash_transactions):
        """
        For samples/Cash Stt _30042018.xlsx, where there is a multi currency
        cash account whose currency field is empty.
        """
        self.assertEqual(len(cash_entries), 3)
        for entry in cash_entries:
            if entry['Currency'] == 'HKD':
                self.assertAlmostEqual(entry['Current Ledger Balance'], 0)
                self.assertAlmostEqual(entry['Current Available Balance'], 0)
                self.assertFalse('Ledger Balance' in entry)

            elif entry['Currency'] == 'USD':
                self.assertAlmostEqual(entry['Ledger Balance'], 38945475.97)
                self.assertAlmostEqual(entry['Current Available Balance'], 7495695.57)
                self.assertTrue('Current Ledger Balance' in entry)

        self.assertEqual(len(cash_transactions), 2)



    def verify_cash_entry1(self, cash_entry):
        """
        The first entry in sample_cash2 _ 16112016.xls
        """
        self.assertEqual(len(cash_entry), 8)
        self.assertEqual(cash_entry['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_entry['Currency'], 'HKD')
        self.assertEqual(cash_entry['Current Ledger Balance'], 0)
        self.assertEqual(cash_entry['Current Available Balance'], 0)



    def verify_cash_entry2(self, cash_entry):
        """
        The last (3rd) entry in sample_cash2 _ 16112016.xls
        """
        self.assertEqual(len(cash_entry), 8)
        self.assertEqual(cash_entry['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_entry['Currency'], 'USD')
        self.assertAlmostEqual(cash_entry['Current Ledger Balance'], 94233724.32)
        self.assertAlmostEqual(cash_entry['Current Available Balance'], 94233724.32)



    def verify_cash_entry01(self, cash_entry):
        """
        The first entry in sample_cash1 _ 16112016.xls
        """
        self.assertEqual(len(cash_entry), 8)
        self.assertEqual(cash_entry['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_entry['Currency'], 'HKD')
        self.assertEqual(cash_entry['Current Ledger Balance'], 0)
        self.assertEqual(cash_entry['Current Available Balance'], 0)



    def verify_cash_entry02(self, cash_entry):
        """
        The last (5th) entry in sample_cash1 _ 16112016.xls
        """
        self.assertEqual(len(cash_entry), 9)
        self.assertEqual(cash_entry['Account Number'], '\'01287508062518')
        self.assertEqual(cash_entry['Account Type'], 'USD Current Account')
        self.assertEqual(cash_entry['Currency'], 'USD')
        self.assertEqual(cash_entry['Ledger Balance'], 62732794.47)
        self.assertEqual(cash_entry['Current Available Balance'], 37598962.52)



    def verify_cash_transaction02(self, cash_transaction):
        """
        from the last (5th) entry in sample_cash1 _ 16112016.xls
        """
        self.assertEqual(len(cash_transaction), 12)
        self.assertEqual(cash_transaction['Account Name'], 'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV')
        self.assertEqual(cash_transaction['Account Number'], '\'01287508062518')
        self.assertEqual(cash_transaction['Currency'], 'USD')
        self.assertEqual(cash_transaction['Processing Date / Time'], datetime(2016,7,11))
        self.assertAlmostEqual(cash_transaction['Amount'], 3176345.83)
        self.assertEqual(cash_transaction['Transaction Reference'], '49969')