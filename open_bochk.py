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


class InvalidHoldingType(Exception):
	pass

class InvalidFieldName(Exception):
	pass

class InconsistentPosition(Exception):
	pass

class InconsistentHolding(Exception):
	pass



def read_bochk(filename, port_values):

	logger.debug('in read_bochk()')

	wb = open_workbook(filename=filename)
	ws = wb.sheet_by_index(0)
	row = 0

	while not field_begins(ws, row):
		row = row + 1

	fields = read_fields(ws, row)

	grand_total = read_holdings(ws, row+1, port_values, fields)
	validate_all_holdings(port_values['holdings'], grand_total)

	logger.debug('out of read_bochk()')



def field_begins(ws, row):
	"""
	Detect whether it has reached the data header row.
	"""
	logger.debug('in field_begins()')
	
	cell_value = ws.cell_value(row, 0)
	if isinstance(cell_value, str) and cell_value.strip() == 'Record Type':
		return True
	else:
		return False



def read_fields(ws, row):
	"""
	ws: the worksheet
	row: the row number to read

	fields = read_fields(ws, row)

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
			logger.error('read_fields(): invalid column name {0} at row {1}, column {2}'.
							format(cell_value, row, column))
			raise InvalidFieldName()

		column = column + 1

	return fields



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
				if value == 'NOM':
					position['settled_units'] = holding_quantity
				elif value == 'PENDING DELIVERY':
					position['pending_delivery'] = position['pending_delivery'] + holding_quantity
				elif value == 'PENDING RECEIPT':
					position['pending_receipt'] = position['pending_receipt'] + holding_quantity
				else:
					logger.error('read_position_holding_detail(): invalid holding type encountered, type={0}, at row {1}'.
									format(position['holding_type'], row))
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
				logger.error('read_position_available_balance(): available balance is not of type float, at row {1}, column {2}'.
								format(row, i))
				raise TypeError

			position['available_balance'] = cell_value

		i = i + 1



def validate_position(position):
	# assume quantity do not have decimal places.
	x = position['settled_units'] - position['pending_delivery'] + \
		position['pending_receipt'] - position['sub_total']

	y = position['settled_units'] - position['pending_delivery'] - \
		position['available_balance']

	if position['quantity_type'] == 'FAMT':
		z = position['sub_total']*position['market_price']/100 - position['market_value']
	else:
		z = position['sub_total']*position['market_price'] - position['market_value']

	if 'exchange_rate' in position:
		z2 = position['market_value']*position['exchange_rate'] - position['equivalent_market_value']
	else:
		z2 = 0


	if x==0 and y==0 and abs(z) < 0.01 and abs(z2) < 0.01:
		pass
	else:
		logger.error('validate_position(): inconsistent position: settled={0}, pending_delivery={1}, pending_receipt={2}, sub_total={3}, available={4}'.
						format(position['settled_units'], position['pending_delivery'],
								position['pending_receipt'], position['sub_total'], position['available_balance']))
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



def is_blank_line(ws, row):
	for i in range(5):
		cell_value = ws.cell_value(row, i)

		if not isinstance(cell_value, str) or not cell_value.strip() == '':
			return False

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
				raise InconsistentPosition()
		else:
			if key in temp_dict:
				merge_position_fields(temp_dict[key], position, key_position_fields)
			else:
				temp_dict[key] = copy_position_fields(position, key_position_fields)

	if not grand_total is None:
		grand_total_position = accumulate_position_total(holdings)
		if abs(grand_total_position - grand_total) > 0.01:
			logger.error('validate_all_holdings(): inconsistent grand_total: position total={0}, grand total={1}'.
							format(grand_total_position, grand_total))
			raise InconsistentPosition()



def is_position_fields_consistent(position1, position2, key_position_fields):
	# Here we assume all holding quantity are of integer value, e.g.
	# quantity 123, 123.0 are fine, but 123.1 are not. Because float
	# comparison is different.
	for fld in key_position_fields:
		if not position1[fld] == position2[fld]:
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