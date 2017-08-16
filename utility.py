# coding=utf-8
# 
# from config_logging package, provides a config object (from config file)
# and a logger object (logging to a file).
# 

import configparser, os



class InvalidDatamode(Exception):
	pass


def get_current_path():
	"""
	Get the absolute path to the directory where this module is in.

	This piece of code comes from:

	http://stackoverflow.com/questions/3430372/how-to-get-full-path-of-current-files-directory-in-python
	"""
	return os.path.dirname(os.path.abspath(__file__))



def _load_config(filename='bochk.config'):
	"""
	Read the config file, convert it to a config object. The config file is 
	supposed to be located in the same directory as the py files, and the
	default name is "config".

	Caution: uncaught exceptions will happen if the config files are missing
	or named incorrectly.
	"""
	path = get_current_path()
	config_file = path + '\\' + filename
	# print(config_file)
	cfg = configparser.ConfigParser()
	cfg.read(config_file)
	return cfg



# initialized only once when this module is first imported by others
if not 'config' in globals():
	config = _load_config()



def get_base_directory():
	"""
	The directory where the log file resides.
	"""
	global config
	directory = config['logging']['directory']
	if directory == '':
		directory = get_current_path()

	return directory



def get_datemode():
	"""
	Read datemode from the config object and return it (in integer)
	"""
	global config
	d = config['excel']['datemode']
	try:
		datemode = int(d)
	except:
		logger.error('get_datemode(): invalid datemode value: {0}'.format(d))
		raise InvalidDatamode()

	return datemode



def get_input_directory():
	"""
	Read directory from the config object and return it.
	"""
	global config
	directory = config['input']['directory']
	if directory.strip() == '':
		directory = get_current_path()

	return directory



def get_exception_list():
	"""
	The list of bonds that allow capital paydown.
	"""
	global config
	return config['bond']['exception_list']