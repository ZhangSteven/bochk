"""
Test the open_jpm.py
"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from bochk.utility import get_current_path
from bochk.open_bochk import read_holdings_bochk, InvalidFieldName, InvalidHoldingType, \
                                InconsistentPosition, InconsistentPositionFieldsTotal, \
                                InconsistentPositionGrandTotal, InvalidCashEntry, \
                                InvalidCashTransaction, read_cash_bochk, read_holdings_bochk, \
                                write_holding_csv, ISINcodeNotFound



class TestBOCHKError(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestBOCHKError, self).__init__(*args, **kwargs)

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


    def test_err(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error.xls'
        port_values = {}
        with self.assertRaises(InvalidFieldName):
            read_holdings_bochk(filename, port_values)



    def test_err2(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error2.xls'
        port_values = {}
        with self.assertRaises(TypeError):
            read_holdings_bochk(filename, port_values)



    def test_err3(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error3.xls'
        port_values = {}
        with self.assertRaises(InvalidHoldingType):
            read_holdings_bochk(filename, port_values)



    def test_err4(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error4.xls'
        port_values = {}
        with self.assertRaises(TypeError):
            read_holdings_bochk(filename, port_values)



    def test_err5(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error5.xls'
        port_values = {}
        with self.assertRaises(TypeError):
            read_holdings_bochk(filename, port_values)



    def test_err6(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error6.xls'
        port_values = {}
        with self.assertRaises(TypeError):
            read_holdings_bochk(filename, port_values)



    def test_err7(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error7.xls'
        port_values = {}
        with self.assertRaises(InconsistentPosition):
            read_holdings_bochk(filename, port_values)



    def test_err8(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error8.xls'
        port_values = {}
        with self.assertRaises(InconsistentPosition):
            read_holdings_bochk(filename, port_values)



    def test_err9(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error9.xls'
        port_values = {}
        with self.assertRaises(InconsistentPositionFieldsTotal):
            read_holdings_bochk(filename, port_values)



    def test_err10(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error10.xls'
        port_values = {}
        with self.assertRaises(InconsistentPositionGrandTotal):
            read_holdings_bochk(filename, port_values)



    def test_cash_error1(self):
        filename = get_current_path() + '\\samples\\cash_error.xls'
        port_values = {}
        with self.assertRaises(InvalidCashEntry):
            read_cash_bochk(filename, port_values)



    def test_cash_error2(self):
        filename = get_current_path() + '\\samples\\cash_error2.xls'
        port_values = {}
        with self.assertRaises(InvalidCashTransaction):
            read_cash_bochk(filename, port_values)



    def test_cash_error3(self):
        filename = get_current_path() + '\\samples\\cash_error3.xls'
        port_values = {}
        with self.assertRaises(InvalidCashTransaction):
            read_cash_bochk(filename, port_values)



    def test_output_error(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error11.xls'
        port_values = {}
        read_holdings_bochk(filename, port_values)
        holding_file = get_current_path() + '\\holding.csv'
    
        with self.assertRaises(ISINcodeNotFound):
            write_holding_csv(holding_file, port_values)