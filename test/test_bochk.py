"""
Test the open_jpm.py
"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from bochk.utility import get_current_path
from bochk.open_bochk import holdings_field_begins, read_holdings_fields, initialize_position, \
                                read_position_holding_detail, read_position_sub_total, \
                                read_position_available_balance, read_position, \
                                validate_position, is_grand_total, read_grand_total, \
                                read_holdings, read_holdings_bochk, retrieve_date_from_filename, \
                                UnhandledFileName, read_cash_bochk



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



    def get_worksheet(self, filename):
        filename = get_current_path() + filename
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_index(0)
        return ws



    def test_retrieve_date_from_filename(self):
        filename = 'C:\\Users\\steven.zhang\\Desktop\\data conversion\\Overseas Bond\\Holding _ 15112016.xls'
        d = retrieve_date_from_filename(filename)
        self.assertEqual(d, datetime(2016,11,15))

        filename = 'C:\\Users\\steven.zhang\\Desktop\\data conversion\\Overseas Bond\\Cash _ 11072016.xls'
        d = retrieve_date_from_filename(filename)
        self.assertEqual(d, datetime(2016,7,11))

        filename = 'C:\\Users\\steven.zhang\\Desktop\\data conversion\\Overseas Bond\\Cash _ 1172016.xls'
        with self.assertRaises(UnhandledFileName):
            retrieve_date_from_filename(filename)



    def test_get_date_from_file(self):
        cash_file = get_current_path() + '\\samples\\Cash _ 15112016.xls'
        holdings_file = get_current_path() + '\\samples\\Holding _ 12072016.xls'
        port_values = {}
        read_cash_bochk(cash_file, port_values)
        read_holdings_bochk(holdings_file, port_values)
        self.assertEqual(datetime(2016,11,15), port_values['cash_date'])
        self.assertEqual(datetime(2016,7,12), port_values['holding_date'])



    def test_holdings_field_begins(self):
        ws = self.get_worksheet('\\samples\\sample_holdings2 _ 16112016.xls')
        row = 0
        self.assertFalse(holdings_field_begins(ws, row))
        row = 2
        self.assertTrue(holdings_field_begins(ws, row))



    def test_read_holdings_fields(self):
        """
        Read the date
        """
        ws = self.get_worksheet('\\samples\\sample_holdings2 _ 16112016.xls')
        row = 0

        while not holdings_field_begins(ws, row):
            row = row + 1

        fields = read_holdings_fields(ws, row)
        self.assertEqual(len(fields), 20)
        self.assertEqual(fields[0], 'generation_date')
        self.assertEqual(fields[10], 'holding_quantity')
        self.assertEqual(fields[19], 'equivalent_market_value')



    def test_read_position_holding_detail(self):
        ws = self.get_worksheet('\\samples\\sample_holdings2 _ 16112016.xls')
        fields = read_holdings_fields(ws, 2)

        position = {}
        initialize_position(position)
        read_position_holding_detail(ws, 3, fields, position)

        self.assertEqual(len(position), 13)
        self.assertEqual(position['generation_date'], datetime(2016,11,16))
        self.assertEqual(position['statement_date'], datetime(2016,11,16))
        self.assertEqual(position['security_id_type'], 'ISIN')
        self.assertEqual(position['security_id'], 'US09681MAC29')
        self.assertEqual(position['settled_units'], 7200000)



    def test_read_position_holding_detail2(self):
        ws = self.get_worksheet('\\samples\\sample_holdings3 _ 16112016.xls')
        fields = read_holdings_fields(ws, 2)

        position = {}
        initialize_position(position)
        read_position_holding_detail(ws, 20, fields, position)
        read_position_holding_detail(ws, 21, fields, position)

        self.assertEqual(len(position), 13)
        self.assertEqual(position['generation_date'], datetime(2016,7,12))
        self.assertEqual(position['statement_date'], datetime(2016,7,11))
        self.assertEqual(position['security_id_type'], 'ISIN')
        self.assertEqual(position['security_id'], 'CNE1000021L3')
        self.assertEqual(position['settled_units'], 0)
        self.assertEqual(position['pending_delivery'], 222000)



    def test_read_position_sub_total(self):
        ws = self.get_worksheet('\\samples\\sample_holdings2 _ 16112016.xls')
        fields = read_holdings_fields(ws, 2)

        # read a normal position
        position = {}
        initialize_position(position)
        read_position_holding_detail(ws, 3, fields, position)
        read_position_sub_total(ws, 4, fields, position)
        self.assertEqual(len(position), 21)
        self.assertEqual(position['sub_total'], 7200000)
        self.assertEqual(position['market_price_currency'], 'USD')
        self.assertAlmostEqual(position['market_price'], 96.966)
        self.assertAlmostEqual(position['market_value'], 6981552)



    def test_read_position_sub_total2(self):
        ws = self.get_worksheet('\\samples\\sample_holdings4 _ 06072016.xls')
        fields = read_holdings_fields(ws, 2)

        # read an All section position
        position = {}
        initialize_position(position)
        read_position_holding_detail(ws, 600, fields, position)
        read_position_sub_total(ws, 601, fields, position)
        self.assertEqual(len(position), 17)
        self.assertEqual(position['sub_total'], 2000000)
        self.assertEqual(position['market_price_currency'], 'USD')
        self.assertAlmostEqual(position['market_price'], 97.4179)
        self.assertAlmostEqual(position['market_value'], 1948358)



    def test_read_position_available_balance(self):
        ws = self.get_worksheet('\\samples\\sample_holdings3 _ 16112016.xls')
        fields = read_holdings_fields(ws, 2)

        position = {}
        read_position_available_balance(ws, 11, fields, position)
        self.assertEqual(position['available_balance'], 3000000)



    def test_read_position(self):
        ws = self.get_worksheet('\\samples\\sample_holdings3 _ 16112016.xls')
        fields = read_holdings_fields(ws, 2)

        position = {}
        row = read_position(ws, 3, fields, position)
        self.assertEqual(row, 6)

        position = {}
        row = read_position(ws, 19, fields, position)
        self.assertEqual(row, 25)

        self.validate_position_fields(position)



    def test_validate_position(self):
        ws = self.get_worksheet('\\samples\\sample_holdings3 _ 16112016.xls')
        fields = read_holdings_fields(ws, 2)

        position = {}
        row = read_position(ws, 3, fields, position)
        try:
            validate_position(position)
        except:
            self.fail('position validation failed')

        position = {}
        row = read_position(ws, 19, fields, position)
        try:
            validate_position(position)
        except:
            self.fail('position 2 validation failed')



    def test_grand_total(self):
        ws = self.get_worksheet('\\samples\\sample_holdings1 _ 10072016.xls')
        fields = read_holdings_fields(ws, 2)

        self.assertTrue(is_grand_total(ws, 93))
        try:
            x = read_grand_total(ws, 93)
            self.assertTrue(x > 0)
        except:
            self.fail('read grand total failed')



    def test_read_holdings(self):
        ws = self.get_worksheet('\\samples\\sample_holdings2 _ 16112016.xls')
        fields = read_holdings_fields(ws, 2)
        port_values = {}
        try:
            x = read_holdings(ws, 3, port_values, fields)
            self.assertTrue(x is None)
            self.verify_holdings(port_values['holdings'])
        except:
            self.fail('read_holdings() failed')



    def test_read_holdings2(self):
        ws = self.get_worksheet('\\samples\\sample_holdings4 _ 06072016.xls')
        fields = read_holdings_fields(ws, 2)
        port_values = {}
        try:
            x = read_holdings(ws, 3, port_values, fields)
            self.assertAlmostEqual(x, 85015806879.97)
            self.verify_holdings2(port_values['holdings'])
        except:
            self.fail('read_holdings() failed')



    def test_read_holdings_bochk(self):
        filename = get_current_path() + '\\samples\\sample_holdings2 _ 16112016.xls'
        port_values = {}
        read_holdings_bochk(filename, port_values)
        self.verify_holdings(port_values['holdings'])



    def test_read_holdings_bochk2(self):
        filename = get_current_path() + '\\samples\\sample_holdings4 _ 06072016.xls'
        port_values = {}
        read_holdings_bochk(filename, port_values)
        self.verify_holdings2(port_values['holdings'])



    # def test_populate_investment_ids(self):
    #     lookup_file = '\\samples\\sample_investmentLookup.xls'
    #     initialize_investment_lookup(lookup_file)
    #     position = {}
    #     position['security_id_type'] = 'ISIN'
    #     position['security_id'] = 'xyz'
    #     portfolio_id = '21815'    # Trading
    #     populate_investment_ids(portfolio_id, position)
    #     self.assertEqual(position['isin'], 'xyz')
    #     self.assertTrue('geneva_investment_id' not in position and \
    #                     'bloomberg_figi' not in position)

    #     position = {}
    #     position['security_id_type'] = 'ISIN'
    #     position['security_id'] = 'xyz'
    #     position['quantity_type'] = 'FAMT'
    #     portfolio_id = '12229'    # HTM
    #     populate_investment_ids(portfolio_id, position)
    #     self.assertEqual(position['geneva_investment_id'], 'xyz HTM')
    #     self.assertTrue('isin' not in position and \
    #                     'bloomberg_figi' not in position)

    #     position = {}
    #     position['security_id_type'] = 'CMU'
    #     position['security_id'] = 'HSBCFN13014'
    #     portfolio_id = '21815'    # Trading
    #     populate_investment_ids(portfolio_id, position)
    #     self.assertEqual(position['isin'], 'HK0000163607')
    #     self.assertEqual(position['bloomberg_figi'], '')
    #     self.assertTrue('geneva_investment_id' not in position)

    #     position = {}
    #     position['security_id_type'] = 'CMU'
    #     position['security_id'] = 'HSBCFN13014'
    #     portfolio_id = '12229'    # HTM
    #     populate_investment_ids(portfolio_id, position)
    #     self.assertEqual(position['geneva_investment_id'], 'HK0000163607 HTM')
    #     self.assertTrue('bloomberg_figi' not in position and \
    #                     'isin' not in position)

    #     position = {}
    #     position['security_id_type'] = 'CMU'
    #     position['security_id'] = 'WLHKFN09007'
    #     portfolio_id = '12229'    # HTM
    #     populate_investment_ids(portfolio_id, position)
    #     self.assertEqual(position['geneva_investment_id'], 'CMU_WLHKFN09007 HTM')
    #     self.assertTrue('bloomberg_figi' not in position and \
    #                     'isin' not in position)



    def validate_position_fields(self, position):
        """
        Fields in a normal position.
        """
        fields = ['generation_date', 'statement_date', 'account_name',
                    'account_number', 'market_code', 'market_name',
                    'security_id_type', 'security_id', 'security_name',
                    'quantity_type', 'market_price_currency',
                    'market_price', 'market_value', 'exchange_currency_pair',
                    'exchange_rate', 'equivalent_currency',
                    'equivalent_market_value', 'settled_units', 'pending_receipt',
                    'pending_delivery', 'sub_total', 'available_balance']

        self.assertEqual(len(position), len(fields))
        for fld in position.keys():
            self.assertTrue(fld in fields)



    def validate_position_fields_All_section(self, position):
        """
        Fields in a position in All section, see samples/sample_holdings4 _ 06072016.xls
        """
        fields = ['generation_date', 'statement_date', 'account_name',
                    'account_number', 'market_code', 'market_name',
                    'security_id_type', 'security_id', 'security_name',
                    'quantity_type', 'market_price_currency',
                    'market_price', 'market_value', 'settled_units', 
                    'pending_receipt', 'pending_delivery', 'sub_total', 
                    'available_balance']
        self.assertEqual(len(position), len(fields))
        for fld in position.keys():
            self.assertTrue(fld in fields)



    def verify_holdings(self, holdings):
        """
        For samples/sample_holdings2 _ 16112016.xls
        """
        self.assertEqual(len(holdings), 27)
        self.validate_position1(holdings[0])
        self.validate_position2(holdings[17])
        self.validate_position3(holdings[26])



    def verify_holdings2(self, holdings):
        """
        For samples/sample_holdings4 _ 06072016.xls
        """
        self.assertEqual(len(holdings), 284)
        self.validate_position01(holdings[0])
        self.validate_position02(holdings[198])
        self.validate_position03(holdings[283])



    def validate_position1(self, position):
        """
        For first position in samples/sample_holdings2 _ 16112016.xls
        """
        self.validate_position_fields(position)
        self.assertEqual(position['statement_date'], datetime(2016,11,16))
        self.assertEqual(position['account_name'], 'MAPLES TRUSTEE S(CY)LTD-CHINA L F TT-CONCORD F INV')
        self.assertEqual(position['security_id'], 'US09681MAC29')
        self.assertEqual(position['sub_total'], 7200000)
        self.assertEqual(position['market_price_currency'], 'USD')
        self.assertAlmostEqual(position['market_price'], 96.966)
        self.assertAlmostEqual(position['market_value'], 6981552)



    def validate_position2(self, position):
        """
        For the 18th position in samples/sample_holdings2 _ 16112016.xls
        """
        self.validate_position_fields(position)
        self.assertEqual(position['statement_date'], datetime(2016,11,16))
        self.assertEqual(position['account_name'], 'MAPLES TRUSTEE S(CY)LTD-CHINA L F TT-CONCORD F INV')
        self.assertEqual(position['security_id'], 'XS1422790615')
        self.assertEqual(position['settled_units'], 7500000)
        self.assertEqual(position['sub_total'], 0)
        self.assertEqual(position['available_balance'], 0)
        self.assertAlmostEqual(position['market_price'], 100.761)



    def validate_position3(self, position):
        """
        For the last (27th) position in samples/sample_holdings2 _ 16112016.xls
        """
        self.validate_position_fields(position)
        self.assertEqual(position['statement_date'], datetime(2016,11,16))
        self.assertEqual(position['security_id'], 'XS1509266026')
        self.assertEqual(position['settled_units'], 5000000)
        self.assertEqual(position['sub_total'], 5000000)
        self.assertEqual(position['available_balance'], 5000000)
        self.assertAlmostEqual(position['market_price'], 99.912)
        self.assertEqual(position['exchange_currency_pair'], 'USD/USD')
        self.assertAlmostEqual(position['equivalent_market_value'], 4995600)



    def validate_position01(self, position):
        """
        For first position in samples/sample_holdings4 _ 06072016.xls
        """
        self.validate_position_fields(position)
        self.assertEqual(position['generation_date'], datetime(2016,7,7))
        self.assertEqual(position['account_name'], 'CLT-CLI HK BR (CLS A-HK)TRUST FUND (SUB-FUND-BOND)')
        self.assertEqual(position['security_id'], 'FR0013101599')
        self.assertEqual(position['sub_total'], 400000)
        self.assertEqual(position['market_price_currency'], 'USD')
        self.assertAlmostEqual(position['market_price'], 107.503)
        self.assertAlmostEqual(position['market_value'], 430012)



    def validate_position02(self, position):
        """
        For the last(199th) holding position in samples/sample_holdings4 _ 06072016.xls
        """
        self.validate_position_fields(position)
        self.assertEqual(position['statement_date'], datetime(2016,7,6))
        self.assertEqual(position['account_name'], 'CLT-CLI OVERSEAS TRUST FD (CAPITAL) (SUB-FD-BOND)')
        self.assertEqual(position['security_id'], 'USY32358AA46')
        self.assertEqual(position['settled_units'], 3000000)
        self.assertEqual(position['sub_total'], 3000000)
        self.assertEqual(position['available_balance'], 3000000)
        self.assertEqual(position['equivalent_currency'], 'HKD')
        self.assertAlmostEqual(position['exchange_rate'], 7.7574)
        self.assertAlmostEqual(position['market_price'], 109.292)



    def validate_position03(self, position):
        """
        For the last (283th) position in All section, in samples/sample_holdings4 _ 06072016.xls
        """
        self.validate_position_fields_All_section(position)
        self.assertEqual(position['statement_date'], datetime(2016,7,6))
        self.assertEqual(position['security_id'], 'US78490FMJ56')
        self.assertEqual(position['market_code'], 'USY')
        self.assertEqual(position['settled_units'], 1000000)
        self.assertEqual(position['sub_total'], 1000000)
        self.assertEqual(position['available_balance'], 1000000)
        self.assertAlmostEqual(position['market_price'], 96.11)
        self.assertEqual(position['market_value'], 961100)