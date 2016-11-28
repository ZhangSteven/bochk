# coding=utf-8
# 
# Read the holdings section of the excel file from trustee.
#
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd
import datetime
from bochk.utility import logger, get_datemode, retrieve_or_create



class InvalidCashTransaction(Exception):
	pass

class InvalidCashEntry(Exception):
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



def read_holdings_bochk(filename, port_values):

	logger.debug('in read_holdings_bochk()')

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



def read_cash_line(ws, row, fields):
	logger.debug('read_cash_line(): at row {0}'.format(row))
	cash_entry = {}
	cash_transaction = None

	column = 0
	for fld in fields:
		cell_value = ws.cell_value(row, column)
		if fld in ['Account Name', 'Account Number', 'Account Type', 'Currency']:
			if isinstance(cell_value, str) and cell_value.strip() != '':
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
		validate_position(position)
		holdings.append(position)

	# end of while loop



def read_position(ws, row, fields, position):
	"""
	Read a position starting at the row number, stops at the
	next position.

	Returns the row number of next line after this position.
	"""
	initialize_position(position)

	i = 0
	cell_value = ws.cell_value(row, 0)
	record_type = cell_value.strip()

	while row < ws.nrows:
		if record_type == 'Holding Details':
			read_position_holding_detail(ws, row, fields, position)
		elif record_type == 'Sub-Total':
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



def read_position_holding_detail(ws, row, fields, position):
	"""
	Read the holding details part of a position.
	"""
	i = 1

	for fld in fields:
		cell_value = ws.cell_value(row, i)
		if fld == 'generation_date' or fld == 'statement_date':
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

	y = position['settled_units'] - position['pending_delivery'] - \
		position['available_balance']

	if position['quantity_type'] == 'FAMT':
		z = position['sub_total']*position['market_price']/100 - position['market_value']
	else:
		z = position['sub_total']*position['market_price'] - position['market_value']

	if 'exchange_rate' in position:	# position in All section has no exchange
		z2 = position['market_value']*position['exchange_rate'] - position['equivalent_market_value']
	else:
		z2 = 0


	if x==0 and y==0 and abs(z) < 0.01 and abs(z2) < 0.01:
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
	cell_value = ws.cell_value(row, 20)
	if isinstance(cell_value, float):
		return cell_value
	else:
		logger.error('read_grand_total(): grand total is not of type float, at row {0}, column 20'.format(row))
		raise TypeError



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
		if abs(grand_total_position - grand_total) > 0.01:
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