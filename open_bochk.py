# coding=utf-8
# 
# Read the holdings section of the excel file from trustee.
#
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import csv, argparse, os, sys
from bochk.utility import logger, get_datemode, get_current_path, \
							get_input_directory



class ISINcodeNotFound(Exception):
	pass

class InvalidPortfolioId(Exception):
	pass

class PortfolioIdNotFound(Exception):
	pass

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



def map_cash_to_portfolio_id(cash_account_name):
	"""
	Map a cash account name to portfolio id.
	"""
	c_map = {
		'MAPLES TRUSTEE SERV (CY) LTD-CHINA LIFE FRANKLIN TT-CONCORD FOCUS INV':'21815',
		'CLT-CLI HK BR (CLASS A- HK TRUST FUND (SUB-FUND-BOND)':'12229',
		'CLT-CLI HK BR (CLASS A- HK) TRUST FUND - SUB FUND I':'12734',
		'CLT-CLI HK BR (CLASS G- HK) TRUST FUND (SUB-FUND-BOND)':'12630',
		'CLT-CLI HK BR TRUST FUND (CAPITAL) (SUB-FUND-BOND)':'12732',
		'CLT-CLI OVERSEAS TRUST FUND (CAPITAL)(SUB-FUND-BOND)':'12733',
		'CLT-CLI MACAU BR (CLASS A-MC) TRUST FUND (SUB-FUND-BOND)':'12366',
		'CLT-CLI HK BR (CLASS A- HK) TRUST FUND (SUB-FUND-TRADING BOND)':'12528',
		'CLT-CHINA LIFE FRANKLIN CLIENTS ACCOUNT':'13456'
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
		'CLT-CLI HK BR (CLS A-HK)TRUST FUND (SUB-FUND-BOND)':'12229',
		'CLT-CLI HK BR(CLASS A-HK) TRUST FUND - SUB FUND I':'12734',
		'CLT-CLI HK BR(CLS G-HK) TRUST FD (SUB-FUND-BOND)':'12630',
		'CLT-CLI HK BR TRUST FUND (CAPITAL) (SUB-FUND-BOND)':'12732',
		'CLT-CLI OVERSEAS TRUST FD (CAPITAL) (SUB-FD-BOND)':'12733',
		'CLT-CLI MACAU BR(CLS A-MC)TRUST FD (SUB-FUND-BOND)':'12366',
		'CLT-CHINA LIFE FRANKLIN CLIENTS ACCOUNT':'13456'
	}

	try:
		return h_map[holding_account_name]
	except KeyError:
		logger.error('map_holding_to_portfolio_id(): {0} is not a valid holding account name'.
						format(holding_account_name))
		raise InvalidHoldingAccountName()



def get_portfolio_accounting_treatment(portfolio_id):
	"""
	Map a portfolio id to its accounting treatment.
	"""
	a_map = {
		'21815':'Trading',
		'12229':'HTM',
		'12734':'HTM',
		'12630':'HTM',
		'12732':'HTM',
		'12733':'HTM',
		'12366':'HTM',
		'13456':'Trading'
	}
	try:
		return a_map[portfolio_id]
	except KeyError:
		logger.error('get_portfolio_accounting_treatment(): {0} is not a valid portfolio id'.
						format(portfolio_id))
		raise InvalidPortfolioId()



def populate_investment_ids(portfolio_id, position):
	"""
	Populate a position with 3 ids:

	1. isin
	2. geneva investment id (only if portfolio is a HTM portfolio and
		the position is a bond)
	3. bloomberg figi (only if no isin is there)
	"""
	if fld == 'geneva_investment_id':

	if position['security_id_type'] == 'ISIN':
		position['isin'] = position['security_id']
	else:
		isin, bbg_id = lookup_isin_code(position['security_id_type'], position['security_id'])
		position['isin'] = isin
		position['bloomberg_figi'] = bbg_id


	# what if a bond:
	# 1. has a CMU code,
	# 2. is in a HTM position,
	#
	# special case handling is needed.
	if (get_portfolio_accounting_treatment(portfolio_id) == 'HTM') and \
		position['quantity_type'] == 'FAMT':

		if position['']




investment_lookup = {}
def initialize_investment_lookup(lookup_file='investmentLookup.xls'):
	"""
	Initialize the lookup table from a file, for those securities that
	do have an isin code.

	To lookup,

	isin, bbg_id = investment_lookup(security_id_type, security_id)
	"""
	filename = get_current_path() + '\\' + lookup_file
	logger.debug('initialize_investment_lookup(): on file {0}'.format(lookup_file))

	wb = open_workbook(filename=filename)
	ws = wb.sheet_by_name('Sheet1')
	row = 1
	global investment_lookup
	while (row < ws.nrows):
		security_id_type = ws.cell_value(row, 0)
		if security_id_type.strip() == '':
			break

		security_id = ws.cell_value(row, 1)
		isin = ws.cell_value(row, 3)
		bbg_id = ws.cell_value(row, 4)
		if isinstance(security_id, float):
			security_id = str(int(security_id))

		investment_lookup[(security_id_type.strip(), security_id.strip())] = \
			(isin.strip(), bbg_id.strip())

		row = row + 1
	# end of while loop 



def lookup_isin_code(security_id_type, security_id):
	global investment_lookup
	if len(investment_lookup) == 0:
		initialize_investment_lookup()

	try:
		return investment_lookup[(security_id_type, security_id)]
	except KeyError:
		logger.error('lookup_isin_code(): No ISIN code found, security_id_type={0}, security_id={1}'.
						format(security_id_type, security_id))
		raise ISINcodeNotFound()



def get_cash_date_as_string(port_values, cash_entry):
	"""
	For BOCHK, there is no date information in the cash file,
	so we lookup the date in the corresponding holdings. In this
	case, we assume the holdings file and the cash file are generated
	on the same day.
	"""
	logger.warning('get_cash_date_as_string(): Using holdings date to represent cash date.')
	holdings = port_values['holdings']
	for position in holdings:
		if map_holding_to_portfolio_id(position['account_name']) == \
						map_cash_to_portfolio_id(cash_entry['Account Name']):
			return convert_datetime_to_string(position['statement_date'])

	logger.error('get_cash_date_as_string(): could not find a portfolio id for cash account:{0}'.
					format(cash_entry['Account Name']))
	raise PortfolioIdNotFound()



def convert_datetime_to_string(dt):
	"""
	convert a datetime object to string in the 'yyyy-mm-dd' format.
	"""
	return '{0}-{1}-{2}'.format(dt.year, dt.month, dt.day)



def write_csv(port_values):
	"""
	Write cash and holdings into csv files.
	"""	
	cash_file = get_input_directory() + '\\cash.csv'
	write_cash_csv(cash_file, port_values)

	holding_file = get_input_directory() + '\\holding.csv'
	write_holding_csv(holding_file, port_values)



def write_cash_csv(cash_file, port_values):
	with open(cash_file, 'w', newline='') as csvfile:
		logger.debug('write_cash_csv(): {0}'.format(cash_file))
		file_writer = csv.writer(csvfile)

		fields = ['Account Number', 'Currency', 'Balance', 'Current Available Balance']
		file_writer.writerow(['Portfolio', 'Date'] + fields)

		for entry in port_values['cash']:
			portfolio_date = get_cash_date_as_string(port_values, entry)
			portfolio_id = map_cash_to_portfolio_id(entry['Account Name'])
			row = [portfolio_id, portfolio_date]

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



def write_holding_csv(holding_file, port_values):
	with open(holding_file, 'w', newline='') as csvfile:
		logger.debug('write_holding_csv(): {0}'.format(holding_file))
		file_writer = csv.writer(csvfile)

		fields = ['statement_date', 'market_code', 'market_name', 'geneva_investment_id',
					'isin', 'bloomberg_figi', 'security_name', 'quantity_type', 
					'settled_units', 'pending_receipt', 'pending_delivery','sub_total',
					'available_balance', 'market_price_currency', 'market_price', 
					'market_value', 'exchange_currency_pair', 'exchange_rate', 
					'equivalent_currency', 'equivalent_market_value']

		file_writer.writerow(['portfolio', 'custodian_account'] + fields)

		for position in port_values['holdings']:
			if position['account_number'] == 'All':
				continue

			portfolio_id = map_holding_to_portfolio_id(position['account_name'])
			custodian_account = 'BOCHK'
			row = [portfolio_id, custodian_account]

			populate_investment_ids(portfolio_id, position)
			for fld in fields:

				if fld == 'statement_date':
					item = convert_datetime_to_string(position[fld])
				else:
					try:
						item = position[fld]
					except KeyError:
						item = ''

				row.append(item)
			# end of inner for

			file_writer.writerow(row)
		# end of outer for loop



if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='Read cash and position files from BOC HK, then convert to Geneva format for reconciliation purpose. Check the config file for path to those files.')
	parser.add_argument('cash_file')
	parser.add_argument('holdings_file')
	args = parser.parse_args()


	cash_file = get_input_directory() + '\\' + args.cash_file
	if not os.path.exists(cash_file):
		print('{0} does not exist'.format(cash_file))
		sys.exit(1)

	holdings_file = get_input_directory() + '\\' + args.holdings_file
	if not os.path.exists(holdings_file):
		print('{0} does not exist'.format(holdings_file))
		sys.exit(1)

	port_values = {}
	try:
		read_cash_bochk(cash_file, port_values)
		read_holdings_bochk(holdings_file, port_values)
		write_csv(port_values)
	except:
		logger.exception('open_bochk:main()')
		print('something goes wrong, check log file.')
	else:
		print('OK')