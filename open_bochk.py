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



class InvalidFieldName(Exception):
	pass



class InconsistentSubtotal(Exception):
	pass



def read_bochk(filename, port_values):
	"""
	
	"""
	logger.debug('in read_bochk()')

	wb = open_workbook(filename=filename)
	ws = wb.sheet_by_index(0)
	row = 0

	while not field_begins(ws, row):
		row = row + 1

	fields = read_fields(ws, row)

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
		'Record Type':'record_type',
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
		'Mnemonic Name':'holding_status',	
		'Market Price Currency':'market_price_currency',
		'Market Unit Price':'market_price',
		'Market Value':'market_value',
		'Exchange Currency Pair':'exchange_currency_pair',
		'Exchange Rate':'exchange_rate',
		'Equivalent Currency':'equivalent_currency',
		'Equivalent Market Value':'equivalent_market_value'
	}

	column = 0
	fields = []

	while column < 21:
		cell_value = ws.cell_value(row, column)
		if isinstance(cell_value, str) and cell_value.strip() in d:
			fields.append(d[cell_value.strip()])
		else:
			logger.error('read_fields(): invalid column name {0} at row {1}, column {2}'.
							format(cell_value, row, column))
			raise InvalidFieldName

		column = column + 1

	return fields