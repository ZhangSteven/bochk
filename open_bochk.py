# coding=utf-8
# 
# Read the holdings section of the excel file from trustee.
#
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import csv, argparse, os, sys, re
from datetime import datetime
from bochk.utility import get_datemode, get_current_path, \
							get_input_directory, get_exception_list
from investment_lookup.id_lookup import get_investment_Ids
import logging
logger = logging.getLogger(__name__)



class UnhandledFileName(Exception):
	pass

class UnhandledPosition(Exception):
	pass

# class PortfolioIdNotFound(Exception):
# 	pass

class InvalidCashAccountName(Exception):
	pass

class InvalidCashTransaction(Exception):
	pass

class InvalidCashEntry(Exception):
	pass

class InvalidHoldingAccountName(Exception):
	pass

class InvalidHoldingType(Exception):
	pass

class InvalidFieldName(Exception):
	pass

class InconsistentPosition(Exception):
	pass

# class InconsistentHolding(Exception):
# 	pass

class InconsistentPositionFieldsTotal(Exception):
	pass

class InconsistentPositionGrandTotal(Exception):
	pass

class FileHandlerNotFound(Exception):
	pass

class GrandTotalNotFound(Exception):
	pass




def read_file(filename, port_values):
	"""
	Read an input file, tell whether it is a holdings file or cash statement
	file, call the holding file or cash file handler to handle it.
	"""	
	fn = filename.split('\\')[-1]	# filename without path
	if fn.startswith('Cash'):
		handler = read_cash_bochk
	elif fn.startswith('Holding'):
		handler = read_holdings_bochk
	elif fn.startswith('BOC Broker Statement'):
		handler = read_holdings_bochk
	elif fn.startswith('BOC Bank Statement'):
		handler = read_cash_bochk
	else:
		logger.error('read_file(): no file handler found for {0}'.format(filename))
		raise FileHandlerNotFound()

	handler(filename, port_values)



def read_holdings_bochk(filename, port_values):

	logger.debug('in read_holdings_bochk()')
	port_values['holding_date'] = retrieve_date_from_filename(filename)

	wb = open_workbook(filename=filename)
	ws = wb.sheet_by_index(0)
	row = 0

	while not holdings_field_begins(ws, row):
		row = row + 1

	fields = read_holdings_fields(ws, row)

	grand_total = read_holdings(ws, row+1, port_values, fields)
	validate_all_holdings(port_values['holdings'], grand_total)

	logger.debug('out of read_holdings_bochk()')



def read_cash_bochk(filename, port_values):

	logger.debug('in read_cash_bochk()')
	port_values['cash_date'] = retrieve_date_from_filename(filename)

	wb = open_workbook(filename=filename)
	ws = wb.sheet_by_index(0)
	row = 0
	
	cell_value = ws.cell_value(row, 0)
	while row < ws.nrows:
		if isinstance(cell_value, str) and cell_value.strip() == 'Account Name':
			break
		row = row + 1

	fields = read_cash_fields(ws, row)
	read_cash(ws, row+1, fields, port_values)

	logger.debug('out of read_cash_bochk()')



def retrieve_date_from_filename(filename):
	"""
	The BOCHK cash and position filenames are of the following format:

	Cash _ ddmmyyyy.xls
	Holding _ ddmmyyyy.xls

	Get the date out of it.
	"""
	fn = filename.split('\\')[-1]	# filename without path
	m = re.search('[0-9]{8}', fn)
	if m != None:
		year = int(m.group(0)[-4:])
		month = int(m.group(0)[2:4])
		day = int(m.group(0)[0:2])
		return datetime(year, month, day)

	else:
		m = re.search('[0-9]{4}-[0-9]{2}-[0-9]{2}', fn)

		if m is None:
			logger.error('retrieve_date_from_filename(): failed to get date from {0}'.
							format(filename))
			raise UnhandledFileName

		year = int(m.group(0)[0:4])
		month = int(m.group(0)[5:7])
		day = int(m.group(0)[8:10])
		return datetime(year, month, day)



def holdings_field_begins(ws, row):
	"""
	Detect whether it has reached the data header row.
	"""
	logger.debug('in holdings_field_begins()')
	
	cell_value = ws.cell_value(row, 0)
	if isinstance(cell_value, str) and cell_value.strip() == 'Record Type':
		return True
	else:
		return False



def read_holdings_fields(ws, row):
	"""
	ws: the worksheet
	row: the row number to read

	fields = read_holdings_fields(ws, row)

	fields: the list of data fields
	"""
	d = {
		'Generation Business Date':'generation_date',
		'Statement Date':'statement_date',	
		'Custody Account Name':'account_name',
		'Custody Account No':'account_number',
		'Market Code':'market_code',
		'Market Name':'market_name',
		'Securities ID Type':'security_id_type',
		'Securities ID':'security_id',
		'Securities name':'security_name',
		'Quantity Type':'quantity_type',
		'Holding':'holding_quantity',
		'Mnemonic Name':'holding_type',	
		'Registration Name':'holding_type',	# the in house fund
		'Market Price Currency':'market_price_currency',
		'Market Unit Price':'market_price',
		'Market Value':'market_value',
		'Exchange Currency Pair':'exchange_currency_pair',
		'Exchange Rate':'exchange_rate',
		'Equivalent Currency':'equivalent_currency',
		'Equivalent Market Value':'equivalent_market_value'
	}

	column = 1
	fields = []

	while column < 21:
		cell_value = ws.cell_value(row, column)
		if isinstance(cell_value, str) and cell_value.strip() in d:
			fields.append(d[cell_value.strip()])
		else:
			logger.error('read_holdings_fields(): invalid column name {0} at row {1}, column {2}'.
							format(cell_value, row, column))
			raise InvalidFieldName()

		column = column + 1

	return fields



def read_cash_fields(ws, row):
	column = 0
	fields = []
	while column < ws.ncols:
		cell_value = ws.cell_value(row, column)
		if isinstance(cell_value, str) and cell_value.strip() == '':
			break

		fields.append(cell_value.strip())
		column = column + 1

	return fields



def read_cash(ws, row, fields, port_values):
	"""
	Read the cash entries.
	"""
	cash = {}
	cash_transactions = []
	port_values['cash_transactions'] = cash_transactions

	while row < ws.nrows:
		if is_blank_line(ws, row):
			break

		cash_entry, cash_tran = read_cash_line(ws, row, fields)
		if not cash_tran is None:
			cash_transactions.append(cash_tran)
		
		update_cash(cash, cash_entry)
		row = row + 1
	# end of while loop

	port_values['cash'] = convert_cash_to_list(cash)



def update_cash(cash, cash_entry):
	key = (cash_entry['Account Name'], cash_entry['Account Number'])
	cash[key] = cash_entry



def convert_cash_to_list(cash):
	cash_list = []
	for key in cash.keys():
		cash_list.append(cash[key])
		# print(cash[key]['Currency'])

	return cash_list



def read_cash_line(ws, row, fields):
	logger.debug('read_cash_line(): at row {0}'.format(row))
	cash_entry = {}
	cash_transaction = None

	column = 0
	for fld in fields:
		cell_value = ws.cell_value(row, column)
		if fld in ['Account Name', 'Account Number', 'Account Type', 'Currency']:
			# if isinstance(cell_value, str) and cell_value.strip() != '':
			if isinstance(cell_value, str):	
				cash_entry[fld] = cell_value.strip()
			else:
				logger.error('read_cash_line(): invalid cash entry at row {0}, column {1}, value={2}'.
								format(row, column, cell_value))
				raise InvalidCashEntry()

		elif fld in ['Hold Amount', 'Float Amount', 'Credit Limit', 
						'Current Ledger Balance', 'Current Available Balance', 
						'Ledger Balance']:
			if isinstance(cell_value, float):
				cash_entry[fld] = cell_value
			else:
				pass	# just leave it

		if fld == 'Processing Date / Time' and not is_empty_cell(cell_value):
			cash_transaction = initialize_cash_transaction(cash_entry)

		if cash_transaction != None and fld in ['Processing Date / Time', 'Value Date']:
			if isinstance(cell_value, float):
				cash_transaction[fld] = xldate_as_datetime(cell_value, get_datemode())
			else:
				logger.error('read_cash_line(): invalid cash transaction at row {0}, column {1}, value={2}'.
								format(row, column, cell_value))
				raise InvalidCashTransaction()

		elif cash_transaction != None and fld == 'Amount':
			if isinstance(cell_value, float):
				cash_transaction[fld] = cell_value
			else:
				logger.error('read_cash_line(): invalid cash transaction at row {0}, column {1}, value={2}'.
								format(row, column, cell_value))
				raise InvalidCashTransaction()

		elif cash_transaction != None and fld in ['Transaction Type', 'Debit / Credit', 'Transaction Reference', 
						'Particulars', 'Cheque Number']:
				if isinstance(cell_value, str):
					cash_transaction[fld] = cell_value.strip()
				elif isinstance(cell_value, float):
					cash_transaction[fld] = str(int(cell_value))

		column = column + 1
	# end of for loop

	return cash_entry, cash_transaction



def initialize_cash_transaction(cash_entry):
	cash_transaction = {}
	for key in cash_entry.keys():
		if key in ['Account Name', 'Account Number', 'Account Type', 'Currency']:
			cash_transaction[key] = cash_entry[key]

	return cash_transaction



def read_holdings(ws, row, port_values, fields):
	"""
	Read the holdings, line by line, until it reaches

	1. The ALL section, if it exists, or,
	2. The end of the holdings.

	Returns the row number where the reading stops, i.e., the next
	line after all the positions.
	"""
	logger.debug('read_holdings(): at row {0}'.format(row))

	holdings = []
	port_values['holdings'] = holdings

	while row < ws.nrows:
		if is_grand_total(ws, row):
			return read_grand_total(ws, row)

		elif is_blank_line(ws, row):
			return None

		position = {}
		row = read_position(ws, row, fields, position)

		# Do not validation positions as occasionally BOCHK makes
		# mistakes in market value computation, e.g., in the 2017-5-5
		# FFX holdings file, first position, the market value should
		# be divided by 100.
		# validate_position(position)

		holdings.append(position)

	# end of while loop



def read_position(ws, row, fields, position):
	"""
	Read a position starting at the row number, stops at the
	next position.

	Returns the row number of next line after this position.
	"""
	logger.debug('read_position(): at row {0}'.format(row))
	initialize_position(position)

	i = 0
	cell_value = ws.cell_value(row, 0)
	record_type = cell_value.strip()

	while row < ws.nrows:
		if record_type == 'Holding Details':
			read_position_holding_detail(ws, row, fields, position)
		elif record_type == 'Sub-Total' or record_type == 'Sub-Total Per Instrument of Custody A/C':
			read_position_sub_total(ws, row, fields, position)
		elif record_type == 'Available Balance':
			read_position_available_balance(ws, row, fields, position)
			break

		row = row + 1
		cell_value = ws.cell_value(row, 0)
		record_type = cell_value.strip()
	# end of while loop

	return row+1



def initialize_position(position):
	position['settled_units'] = 0
	position['pending_receipt'] = 0
	position['pending_delivery'] = 0
	position['pending_call'] = 0



def read_position_holding_detail(ws, row, fields, position):
	"""
	Read the holding details part of a position.
	"""
	i = 1

	for fld in fields:
		cell_value = ws.cell_value(row, i)
		if fld == 'generation_date' or fld == 'statement_date':
			# print(cell_value)
			position[fld] = xldate_as_datetime(cell_value, get_datemode())
		elif fld == 'holding_quantity':
			if isinstance(cell_value, float):
				holding_quantity = cell_value
			else:
				logger.error('read_position_holding_detail(): holding_quantity is not of type float, at row {0}, column {1}'.
								format(row, i))
				raise TypeError

		else:
			value = cell_value.strip()
			if fld == 'holding_type':
				# print('holding type: {0}'.format(value))
				if value == 'NOM':
					position['settled_units'] = holding_quantity
				elif value == 'ENT':	# the position is going to be called
					position['settled_units'] = holding_quantity
					position['pending_call'] = holding_quantity
				elif value == 'PENDING DELIVERY':
					position['pending_delivery'] = position['pending_delivery'] + holding_quantity
				elif value == 'PENDING RECEIPT':
					position['pending_receipt'] = position['pending_receipt'] + holding_quantity
				else:
					logger.error('read_position_holding_detail(): invalid holding type encountered, type={0}, at row {1}'.
									format(value, row))
					raise InvalidHoldingType()

				break

			else:
				position[fld] = value

		i = i + 1



def read_position_sub_total(ws, row, fields, position):
	"""
	Read the sub-total part of a position
	"""
	i = 1

	for fld in fields:
		cell_value = ws.cell_value(row, i)

		if fld in ['market_price_currency', 'exchange_currency_pair', 'equivalent_currency']:
			position[fld] = cell_value.strip()
		elif fld in ['holding_quantity', 'market_price', 'market_value', 'exchange_rate', 'equivalent_market_value']:
			if not isinstance(cell_value, float):
				logger.error('read_position_sub_total(): {0} is not of type float, at row {1}, column {2}'.
								format(fld, row, i))
				raise TypeError

			if fld == 'holding_quantity':
				position['sub_total'] = cell_value
			else:
				position[fld] = cell_value

		if position['account_number'] == 'All' and fld == 'market_value':
			break

		i = i + 1



def read_position_available_balance(ws, row, fields, position):
	"""
	Read the available balance part of the position.
	"""
	i = 1
	for fld in fields:
		cell_value = ws.cell_value(row, i)
		if fld == 'holding_quantity':
			if not isinstance(cell_value, float):
				logger.error('read_position_available_balance(): available balance is not of type float, at row {0}, column {1}'.
								format(row, i))
				raise TypeError

			position['available_balance'] = cell_value

		i = i + 1



def validate_position(position):
	"""
	Make sure the position's quantity, market price, market value are
	consistent.

	Assume quantity do not have decimal places.
	"""
	x = position['settled_units'] - position['pending_delivery'] + \
		position['pending_receipt'] - position['sub_total']

	# y = position['settled_units'] - position['pending_delivery'] - \
	# 	position['pending_call'] - position['available_balance']

	if position['quantity_type'] == 'FAMT':
		z = position['sub_total']*position['market_price']/100 - position['market_value']
	else:
		z = position['sub_total']*position['market_price'] - position['market_value']

	if 'exchange_rate' in position:	# position in All section has no exchange
		z2 = position['market_value']*position['exchange_rate'] - position['equivalent_market_value']
	else:
		z2 = 0

	if position['sub_total'] < 0:	# for short positions, available balance = 0
		# y = 0						# market value of position is also 0.
		z = 0
		z2 = 0

	# if x==0 and y==0 and abs(z) < 0.01 and abs(z2) < 0.01:
	if x==0 and abs(z) < 0.01 and abs(z2) < 0.01:
		pass
	elif position['security_id_type']+':'+position['security_id'] in get_exception_list():
		# if it is a bond (ABS etc.) that allows capital paydown so that market
		# value is not the product of quantity and price.
		pass
	else:
		logger.error('validate_position(): inconsistent position: market={0}, {1}={2}, settled_units={3}'.
						format(position['market_code'], position['security_id_type'],
								position['security_id'], position['settled_units']))
		raise InconsistentPosition()



def is_grand_total(ws, row):
	cell_value = ws.cell_value(row, 0)
	if cell_value.strip() == 'Grand Total of All Custody A/C':
		return True
	else:
		return False



def read_grand_total(ws, row):
	# cell_value = ws.cell_value(row, 20)
	# if isinstance(cell_value, float):
	# 	return cell_value
	# else:
	# 	logger.error('read_grand_total(): grand total is not of type float, at row {0}, column 20'.format(row))
	# 	raise TypeError

	for i in range(1, ws.ncols):
		cell_value = ws.cell_value(row, i)
		if isinstance(cell_value, str) and cell_value.strip().lower() in ['hkd', 'usd']:

			if isinstance(ws.cell_value(row, i+1), float):
				return ws.cell_value(row, i+1)
			else:
				logger.error('read_grand_total(): grand total {0} should be float, at row {1} column {2}'.
								format(ws.cell_value(row, i+1), row, i+1))
				raise TypeError

	logger.error('read_grand_total(): grand total not found')
	raise GrandTotalNotFound()



def is_blank_line(ws, row):
	for i in range(5):
		cell_value = ws.cell_value(row, i)
		if not is_empty_cell(cell_value):
			return False

	return True



def is_empty_cell(cell_value):
	if not isinstance(cell_value, str) or cell_value.strip() != '':
		return False
	else:
		return True


def validate_all_holdings(holdings, grand_total):
	temp_dict = {}
	key_position_fields = ['settled_units', 'pending_receipt', 'pending_delivery', 
							'sub_total', 'available_balance']

	for position in holdings:
		key = (position['market_code'], position['security_id_type'], position['security_id'])
		if position['account_number'] == 'All':
			logger.debug('validate_all_holdings(): validate security {0}'.format(position['security_id']))
			if not is_position_fields_consistent(temp_dict[key], position, key_position_fields):
				logger.error('validate_all_holdings(): inconsistent positions: {0}'.format(key))
				raise InconsistentPositionFieldsTotal()
		else:
			if key in temp_dict:
				merge_position_fields(temp_dict[key], position, key_position_fields)
			else:
				temp_dict[key] = copy_position_fields(position, key_position_fields)

	if not grand_total is None:
		grand_total_position = accumulate_position_total(holdings)
		if abs(grand_total_position - grand_total) > 0.5:	# the in house fund seems to
															# keep one decimal point only,
															# therefore we use 0.5 as the
															# threshoold
			logger.error('validate_all_holdings(): inconsistent grand total: position total={0}, grand total={1}'.
							format(grand_total_position, grand_total))
			raise InconsistentPositionGrandTotal()



def is_position_fields_consistent(position1, position2, key_position_fields):
	# Here we assume all holding quantity are of integer value, e.g.
	# quantity 123, 123.0 are fine, but 123.1 are not. Because float
	# comparison is different.
	for fld in key_position_fields:
		if position1[fld] != position2[fld]:
			return False

	return True



def merge_position_fields(position1, position2, key_position_fields):
	for fld in key_position_fields:
		position1[fld] = position1[fld] + position2[fld]



def copy_position_fields(position, key_position_fields):
	new_position = {}
	for fld in key_position_fields:
		new_position[fld] = position[fld]

	return new_position
	


def accumulate_position_total(holdings):
	t = 0
	for position in holdings:
		if position['account_number'] != 'All':
			t = t + position['equivalent_market_value']

	return t



def map_cash_to_portfolio_id(cash_account_name):
	"""
	Map a cash account name to portfolio id.
	"""
	c_map = {
		'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV':'21815',
		
		# the old and new cash account name for 12229. The old mapping
		# (the first one) are kept so that old tests won't break.
		'CLT-CLI HK BR (CLASS A- HK TRUST FUND (SUB-FUND-BOND)':'12229',
		'CLT - CLI HK BR (CLASS A-HK) TRUST FUND (BOND)- PAR':'12229',

		# the old and new cash account name for 12734
		'CLT-CLI HK BR (CLASS A- HK) TRUST FUND - SUB FUND I':'12734',
		'CLT - CLI HK BR (CLASS A-HK) TRUST FUND (BOND)':'12734',

		# old and new names for 12630
		'CLT-CLI HK BR (CLASS G- HK) TRUST FUND (SUB-FUND-BOND)':'12630',
		'CLT - CLI HK BR (CLASS G-HK) TRUST FUND (BOND)':'12630',

		'CLT-CLI HK BR TRUST FUND (CAPITAL) (SUB-FUND-BOND)':'12732',
		'CLT-CLI OVERSEAS TRUST FUND (CAPITAL)(SUB-FUND-BOND)':'12733',

		# old and new names for 12366
		'CLT-CLI MACAU BR (CLASS A-MC) TRUST FUND (SUB-FUND-BOND)':'12366',
		'CLT - CLI MACAU BR (CLASS A-MC) TRUST FUND (BOND)':'12366',

		# account 12298
		'CLT-CLI MACAU BR (CLASS A-MC) TRUST FUND': '12298',

		# old and new names for 12528
		'CLT-CLI HK BR (CLASS A- HK) TRUST FUND (SUB-FUND-TRADING BOND)':'12528',
		'CLT - CLI HK BR (CLASS A-HK) TRUST FUND':'12528',

		'CLT - CLI MACAU BR (CLASS A-MC) TRUST FUND - PAR': '13006',

		'CLT - CLI MACAU BR (CLASS A-MC) TRUST FUND (BOND) - PAR':'12549',

		'CLT-CHINA LIFE FRANKLIN CLIENTS ACCOUNT':'13456',
		'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-GREEN BLUE SP OP F':'11602',
		'MAPLES TRUSTEE SERVICE(CY)LTD-CHINA LIFE FRANKLIN TT-SPECIAL EVENT FD':'16454',
		'MAPLES TRUSTEE SERV(CY) LTD - CHINA LIFE FRANKLIN TT - FFX INVESTMENTS':'30001',

		# the cash of 11490 in BOCHK, old and new names
		'CLT-CLI HK BR (CLASS A- HK) TRUST FUND':'11490',
		'CLT-CLI HK BR (Class A-HK) - Par TRUST FUND':'11490',
		'CLT - CLI HK BR (CLASS A-HK) TRUST FUND - PAR': '11490',

		# the in house fund
		'CHINA LIFE FRANKLIN ASSET MANAGEMENT CO LTD':'20051',

		# JIC International
		'JIC INTERNATIONAL LIMITED - CLFAMC': '40002',

		# FIXME: All client accounts share the same account name,
		# 
		'CHINA LIFE FRANKLIN ASSET MANAGEMENT CO.': '40004'
	}

	try:
		return c_map[cash_account_name]
	except KeyError:
		logger.error('map_cash_to_portfolio_id(): {0} is not a valid cash account name'.
						format(cash_account_name))
		raise InvalidCashAccountName()



def map_holding_to_portfolio_id(holding_account_name):
	"""
	Map a holding position account name to portfolio id.
	"""
	h_map = {
		'MAPLES TRUSTEE S(CY)LTD-CHINA L F TT-CONCORD F INV':'21815',

		# old and new names of 12229
		'CLT-CLI HK BR (CLS A-HK)TRUST FUND (SUB-FUND-BOND)':'12229',
		'CLT - CLI HK BR (CLASS A-HK) TRUST FD (BOND)- PAR':'12229',

		# old and new names of 12734
		'CLT-CLI HK BR(CLASS A-HK) TRUST FUND - SUB FUND I':'12734',
		'CLT - CLI HK BR (CLASS A-HK) TRUST FUND (BOND)':'12734',

		# old and new names of 12630
		'CLT-CLI HK BR(CLS G-HK) TRUST FD (SUB-FUND-BOND)':'12630',
		'CLT - CLI HK BR (CLASS G-HK) TRUST FUND (BOND)':'12630',

		'CLT-CLI HK BR TRUST FUND (CAPITAL) (SUB-FUND-BOND)':'12732',
		'CLT-CLI OVERSEAS TRUST FD (CAPITAL) (SUB-FD-BOND)':'12733',

		# old and new names for 12366
		'CLT-CLI MACAU BR(CLS A-MC)TRUST FD (SUB-FUND-BOND)':'12366',
		'CLT - CLI MACAU BR (CLASS A-MC) TRUST FUND (BOND)':'12366',

		'CLT-CLI HK BR(CLS A-HK)TRUST FD(SUB-FD-TRADING BD)':'12528',
		'CLT-CHINA LIFE FRANKLIN CLIENTS ACCOUNT':'13456',
		'MAPLES T SER(CY)LTD-CL FRANK TT-GREEN BLUE SP OP F':'11602',
		'MAPLES TRUSTEE SERV (CY) LTD-CHINA L F TT-S E FD':'16454',
		'MAPLES TRUSTEE SERV(CY) LTD-CHINA L F TT-FFX INV':'30001',

		
		'CLT-CLI MACAU BR (CLASS A-MC) TRUST FUND': '12298',

		'CLT - CLI MACAU BR (CLASS A-MC) TRUST FUND - PAR': '13006',

		# the in house fund
		'CHINA LIFE FRANKLIN ASSET MANAGEMENT CO LTD':'20051',

		# the new fund from China Life Macau (temporary port code)
		'CLT-CLI MACAU BR(CLS G-MC)TRUST FD (SUB-FUND-BOND)':'99999',

		'CLT - CLI MACAU BR (CLASS A-MC) TRUST FD (BD)-PAR':'12549',

		# JIC International
		'JIC INTERNATIONAL LIMITED - CLFAMC': '40002',

		# Client Account Starberry
		'CHINA LIFE FRANKLIN ASSET MGT CO LTD-CLIENT A/C 2': '40004',

		'CLT - CLI HK BR (CLASS A-HK) TRUST FUND - PAR': '11500'
	}

	try:
		return h_map[holding_account_name]
	except KeyError:
		logger.error('map_holding_to_portfolio_id(): {0} is not a valid holding account name'.
						format(holding_account_name))
		raise InvalidHoldingAccountName()



def convert_datetime_to_string(dt):
	"""
	convert a datetime object to string in the 'yyyy-mm-dd' format.
	"""
	return '{0}-{1}-{2}'.format(dt.year, dt.month, dt.day)



def get_prefix_from_dir(input_dir):
	"""
	Work out a prefix for the filename depending on the input directory.
	"""
	folder_name = input_dir.split('\\')[-1]
	prefix = ''
	for token in folder_name.lower().split():
		prefix = prefix + token + '_'

	return prefix + 'bochk_'



def create_csv_file_name(date, output_dir, file_prefix, file_suffix):
	"""
	Create the output csv file name based on the date string, as well as
	the file suffix: cash, afs_positions, or htm_positions
	"""
	date_string = convert_datetime_to_string(date)
	csv_file = output_dir + '\\' + file_prefix + date_string + '_' \
				+ file_suffix + '.csv'
	return csv_file



def write_cash_or_holding_csv(port_values, directory=get_input_directory(),
								file_prefix=get_prefix_from_dir(get_input_directory())):
	"""
	Write cash or holdings into csv files.
	"""
	output_file = write_cash_csv(port_values, directory, file_prefix)
	if output_file is None:	# it's a holding file instead of cash file
		output_file = write_holding_csv(port_values, directory, file_prefix)

	return output_file



def write_csv(port_values, directory=get_input_directory(),
				file_prefix=get_prefix_from_dir(get_input_directory())):
	"""
	Write cash and holdings into csv files.
	"""	
	write_cash_csv(port_values, directory, file_prefix)
	write_holding_csv(port_values, directory, file_prefix)



def consolidate_cash(port_values):
	"""
	Combine the checking and savings account for the same currency in
	the same bank.
	"""
	new_cash_accounts = []
	cash_accounts = port_values['cash']
	for cash_account in cash_accounts:
		if find_n_merge(cash_account, new_cash_accounts):
			continue

		if cash_account['Currency'] == '':
			continue

		new_cash_accounts.append(cash_account)

	port_values['cash'] = new_cash_accounts

	

def find_n_merge(cash_account, cash_accounts):
	"""
	find accounts under the same fund, with the same currency, e.g., savings
	account and current account of the same currency, merge them into one.
	"""
	for ca in cash_accounts:
		if cash_account['Currency'] == ca['Currency'] and \
			cash_account['Account Name'] == ca['Account Name']:

			try:
				ca['Current Ledger Balance'] = ca['Current Ledger Balance'] + cash_account['Current Ledger Balance']
			except KeyError:
				pass
		
			try:
				ca['Current Available Balance'] = ca['Current Available Balance'] + cash_account['Current Available Balance']
			except KeyError:
				pass

			try:
				ca['Ledger Balance'] = ca['Ledger Balance'] + cash_account['Ledger Balance']
			except KeyError:
				pass

			return True

	return False



def write_cash_csv(port_values, directory, file_prefix):
	if not 'cash' in port_values:	# do nothing
		logger.warning('write_cash_csv(): no cash information is found.')
		return None

	cash_file = create_csv_file_name(port_values['cash_date'], directory, file_prefix, 'cash')
	with open(cash_file, 'w', newline='') as csvfile:
		logger.debug('write_cash_csv(): {0}'.format(cash_file))
		file_writer = csv.writer(csvfile, delimiter='|')

		fields = ['Account Number', 'Currency', 'Balance', 'Current Available Balance']
		file_writer.writerow(['Portfolio', 'Date', 'Custodian'] + fields)

		consolidate_cash(port_values)
		for entry in port_values['cash']:
			# portfolio_date = get_cash_date_as_string(port_values, entry)
			cash_date = convert_datetime_to_string(port_values['cash_date'])
			portfolio_id = map_cash_to_portfolio_id(entry['Account Name'])
			row = [portfolio_id, cash_date, 'BOCHK']

			for fld in fields:
				if fld == 'Balance':
					try:
						item = entry['Ledger Balance']
					except KeyError:
						item = entry['Current Ledger Balance']
				else:
					item = entry[fld]

				row.append(item)
    		# end of for loop

			file_writer.writerow(row)

	return cash_file



def write_holding_csv(port_values, directory, file_prefix):
	if not 'holdings' in port_values:	# do nothing
		logger.warning('write_holding_csv(): no holding information is found.')
		return None

	holding_file = create_csv_file_name(port_values['holding_date'], directory, file_prefix, 'position')
	with open(holding_file, 'w', newline='') as csvfile:
		logger.debug('write_holding_csv(): {0}'.format(holding_file))
		file_writer = csv.writer(csvfile, delimiter='|')

		fields = ['market_code', 'market_name', 'security_name', 
					'quantity_type', 'settled_units', 'pending_receipt', 
					'pending_delivery', 'sub_total', 'available_balance', 
					'market_price_currency', 'market_price', 'market_value', 
					'exchange_currency_pair', 'exchange_rate', 
					'equivalent_currency', 'equivalent_market_value', 'pending_call']

		file_writer.writerow(['portfolio', 'custodian_account', 'geneva_investment_id',
								'isin', 'bloomberg_figi', 'date'] + fields)

		for position in port_values['holdings']:
			if position['account_number'] == 'All':
				continue

			portfolio_id = map_holding_to_portfolio_id(position['account_name'])
			row = [portfolio_id, 'BOCHK']

			investment_ids = get_investment_Ids(portfolio_id, position['security_id_type'], 
												position['security_id'])
			for id in investment_ids:
				row.append(id)

			row.append(convert_datetime_to_string(port_values['holding_date']))

			for fld in fields:

				try:
					item = position[fld]
				except KeyError:
					item = ''

				row.append(item)
			# end of inner for

			file_writer.writerow(row)
		# end of outer for loop
	return holding_file



if __name__ == '__main__':
	"""
	Generate cash or holdings csv files on demand, use

	python open_bochk.py <holding_or_cash_file>

	Or,

	python open_bochk.py <holding_file> <cash_file>

	the order of the holding file and cash doesn't matter.

	Note if you put 2 holding files or 2 cash files, it won't give you
	an error, but the output of the latter will override the output csv
	of the former.
	"""
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)
	
	parser = argparse.ArgumentParser(description='Read cash and position files from BOC HK, then convert to Geneva format for reconciliation purpose. Check the config file for path to those files.')
	parser.add_argument('files', metavar='files', type=str, nargs='+',
						help='cash and/or holdings files')
	args = parser.parse_args()

	for file in args.files:
		file = get_input_directory() + '\\' + file

		if not os.path.exists(file):
			print('{0} does not exist'.format(file))
			sys.exit(1)

		try:
			port_values = {}
			read_file(file, port_values)
			write_cash_or_holding_csv(port_values)
			print('OK')
		except:
			logger.exception('open_bochk:main()')
			print('something goes wrong, check log file.')
			sys.exit(1)
