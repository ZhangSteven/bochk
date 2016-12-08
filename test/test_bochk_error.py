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
                                write_holding_csv, write_csv, UnhandledPosition, \
                                InvalidCashAccountName
from investment_lookup.id_lookup import InvestmentIdNotFound




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
        filename = get_current_path() + '\\samples\\sample_holdings_error _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InvalidFieldName):
            read_holdings_bochk(filename, port_values)



    def test_err2(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error2 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(TypeError):
            read_holdings_bochk(filename, port_values)



    def test_err3(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error3 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InvalidHoldingType):
            read_holdings_bochk(filename, port_values)



    def test_err4(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error4 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(TypeError):
            read_holdings_bochk(filename, port_values)



    def test_err5(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error5 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(TypeError):
            read_holdings_bochk(filename, port_values)



    def test_err6(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error6 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(TypeError):
            read_holdings_bochk(filename, port_values)



    def test_err7(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error7 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InconsistentPosition):
            read_holdings_bochk(filename, port_values)



    def test_err8(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error8 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InconsistentPosition):
            read_holdings_bochk(filename, port_values)



    def test_err9(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error9 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InconsistentPositionFieldsTotal):
            read_holdings_bochk(filename, port_values)



    def test_err10(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error10 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InconsistentPositionGrandTotal):
            read_holdings_bochk(filename, port_values)



    def test_cash_error1(self):
        filename = get_current_path() + '\\samples\\cash_error _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InvalidCashEntry):
            read_cash_bochk(filename, port_values)



    def test_cash_error2(self):
        filename = get_current_path() + '\\samples\\cash_error2 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InvalidCashTransaction):
            read_cash_bochk(filename, port_values)



    def test_cash_error3(self):
        filename = get_current_path() + '\\samples\\cash_error3 _ 16112016.xls'
        port_values = {}
        with self.assertRaises(InvalidCashTransaction):
            read_cash_bochk(filename, port_values)



    def test_cash_error4(self):
        holdings_file = get_current_path() + '\\samples\\sample_holdings2 _ 16112016.xls'
        cash_file = get_current_path() + '\\samples\\cash_error4 _ 16112016.xls'
        port_values = {}
        directory = get_current_path() + '\\samples'
        read_cash_bochk(cash_file, port_values)
        read_holdings_bochk(holdings_file, port_values)
        with self.assertRaises(InvalidCashAccountName):
            write_csv(port_values, directory)



    # def test_populate_investment_ids(self):
    #     lookup_file = '\\samples\\sample_investmentLookup.xls'
    #     initialize_investment_lookup(lookup_file)
    #     position = {}
    #     position['security_id_type'] = 'ISIN'
    #     position['security_id'] = 'xyz'
    #     position['quantity_type'] = 'units' # not a bond
    #     portfolio_id = '12229'    # HTM
    #     with self.assertRaises(UnhandledPosition):
    #         populate_investment_ids(portfolio_id, position)



    def test_output_error(self):
        filename = get_current_path() + '\\samples\\sample_holdings_error11 _ 16112016.xls'
        port_values = {}
        read_holdings_bochk(filename, port_values)
        directory = get_current_path() + '\\samples'
    
        with self.assertRaises(InvestmentIdNotFound):
            write_holding_csv(port_values, directory)