# bochk

This is to convert the BOCHK broker statement into files containing investment positions and cash to reconcile with Advent Geneva system.

The positions are trade day positions, cash is settlement day cash.


++++++++++
Note
++++++++++
1. The program assumes "statement date" in the holdings file as the date for the positions and cash, is that true? Should we use the date in the filename?

2. The program hard coded some of the holding account name and cash name, to be mapped to portfolio ids, look out when new portfolios are added.

3. The program assumes a portfolio holds either all HTM or all Trading bonds, see lookup_accounting_treatment().

4. When a position's security id type is not ISIN, the program then looks up the isin code or BBG ID in investmentLookup.xls file. Currently there are 2 such bonds (based on 2016-7-6 broker statement), when more such bonds appear, update this file.



++++++++++
ver 0.34@2018-5-17
++++++++++
1. Add support for cash statements containing multiple currency savings account whose currency is empty and balance is zero. This type of account is simply ignored.



++++++++++
ver 0.33@2017-10-24
++++++++++
1. Fixed bug in find_n_merge() in open_bochk.py, where it combines accounts from different fund of the same currency into one account. Now it only combines accounts from the same fund, of the same currency, into one account.



++++++++++
ver 0.32@2017-10-20
++++++++++
1. Added support for in house fund, also from BOCHK.
2. In house fund has a slight different column name "Registration Name" in holdings excel, and multiple cash accounts of the same currency in cash excel. Therefore cash consolidation and testing code are updated, two more samples are added, too.



++++++++++
ver 0.31@2017-8-16
++++++++++
1. Updated logging, now all modules use the standard way to obtain a logger:
	logger = logging.getLogger(__name__)



++++++++++
ver 0.302
++++++++++
1. Since validate_position() function is no longer called, then error testing code also updated to reflect that change. test_bochk_error.py, test_err8()



++++++++++
ver 0.301
++++++++++
1. The validate_position() function is no longer called in read_holdings() function, because BOCHK occasionally makes mistakes in market value computation. For example, FFX fund, 2017-5-5 holdings file, fist position. This is not considerred an upgrade of the program, rather, a comproise to make it work. Therefore the version number 0.301.



++++++++++
ver 0.30
++++++++++
1. The read_grand_total() function is modified, previously it reads grand total from column 20, but we find that some times it is not there. The function now searches for the value on that row.



++++++++++
ver 0.29
++++++++++
1. When there is a short position, position sub total is negative, but available balance is zero. This fails the validation. The code is modified to suit this case.



++++++++++
ver 0.28
++++++++++
1. Move the new column added in ver 0.26 to the last column, because it was inserted in between other columns and that caused the custom loader in Geneva to load wrong data, because that loader map column number to data.



++++++++++
ver 0.27
++++++++++
1. Add mapping from fund name to portfolio code (both cash and position), for special event fund.



++++++++++
ver 0.26
++++++++++
1. Added support for bond call. When bond call is announced, the holding type is 'ENT', meaning the bond is still available for sell, so sub total won't change, but available balance drops.




++++++++++
ver 0.25
++++++++++
1. The config file is changed to add an option allowing the validate_position() function to ignore certain bonds, because they allow principle paydown so they won't satisfy the validation checks and they need to be ignored for the check.



++++++++++
ver 0.2402
++++++++++
1. It's found that "CLT-CLI HK BR (CLASS A- HK) TRUST FUND" in BOCHK is actually 11490's cash holding in BOCHK. So we fixed that in the map_cash_to_portfolio_id() function.



++++++++++
ver 0.2401
++++++++++
1. Add fund name mapping "CLT-CLI HK BR (CLASS A- HK) TRUST FUND" to portfolio code '99999'. Please find out what out it actually is.



++++++++++
ver 0.24
++++++++++
1. Add fund name to portfolio id mapping for Greenblue.



++++++++++
ver 0.2301
++++++++++
1. Add one more test case for get date from file name, for file names like BOC Bank Statement 2016-01-29-30 (CLASS A-MC SUB FUND BOND) -HKD.xls



++++++++++
ver 0.23
++++++++++
1. Bug fix: when both a cash and holdings file are present on the command line, only the first file got output.



++++++++++
ver 0.22
++++++++++
1. Add read_file() function, it tells whether the input file is a holdings file or cash file. It then calls the read_holdings_bochk() or read_cash_bochk() functions accordingly.

2. Change the write_cash_csv() and write_holding_csv() function so that they will ignore the input if the port_values does not contain cash or holdings information.



++++++++++
ver 0.21
++++++++++
1. Change the write_cash_csv() and write_holding_csv() function so that they return the output csv filename (full path). This is required by the recon_helper package.



++++++++++
ver 0.20
++++++++++
1. Change the write_csv(), write_cash_csv() and write_holding_csv() function, now they take an additional argument as filename prefix, to work with reconciliaiton_helper package.



++++++++++
ver 0.1901
++++++++++
1. Change the configure file, so that input directory is for office PC, previously it was for hong kong home laptop.



++++++++++
ver 0.19
++++++++++
1. Now output csv file name solely depends on the input directory folder name, and it will always contain "bochk", e.g, input path is C:\...\concord, then output csv file will be "concord_bochk_*.csv".



++++++++++
ver 0.18
++++++++++
1. Now output csv file name depends on the input directory, the mapping is as follows (directory case insensitive):

	directory 	: file name prefix
	
	Concord		: 21815_*.csv
	Greenblue	: 11602_*.csv
	CLO bond    : clo_bond_*.csv
	Special Event Fund: 16454_*.csv
	in-house Fund : 88888_*.csv



++++++++++
ver 0.17
++++++++++
1. Bug fix: the date of the position should be extracted from the file name instead of the 'statement_date' in the positions. Because the latter means the date when the statement has been generated.

2. Change the output csv to use '|' as delimiter, to avoid potential problem due to data field such as "security name" containing commas.

3. Add date to the output csv file name.



++++++++++
ver 0.16
++++++++++
1. Move investment id lookup and portfolio accounting lookup functions to another project investment_lookup, so that we have centralized control on these settings.



++++++++++
ver 0.15
++++++++++
1. Bug fix: in utility.py, the get_datemode() function raises nothing when datemode value is invalid.



++++++++++
ver 0.14
++++++++++
1. Bug fix: when the configure file changes the path (we have different input file path in office and home PCs), one test fails.

2. The investmentLookup.xls is added more comments.



++++++++++
ver 0.13
++++++++++
1. Bug fix: when a position's security_id_type is not "ISIN", but can actually lookup an isin, its geneva investmend id fo HTM position now is isin + " HTM".



++++++++++
ver 0.12
++++++++++
1. In output csv file, the geneva_investment_id column replaces the accounting_treatment column, to solve the problem of HTM bonds may not having ISIN code in Geneva security master. For a HTM position, only the geneva_investment_id column will have output, where for trading position, either the isin or bloomberg_figi column have output.

2. In the investmentLookup.xls file, an extra column of geneva_investment_id for HTM position is also added.



++++++++++
ver 0.11
++++++++++
1. Add two entries in the config file:

	> base directory for input cash/position files and output the csv files. So those files do not mix with the code.

	> base directory for the log file. So during production deployment, the log file can be put in a different directory for easy checking.

2. logging function is handled by another package config_logging.

3. Use argparse to replace the original sys.args approach.


++++++++++
ver 0.1
++++++++++
1. Generates a cash and holdings csv file from BOCHK's holdings and cash statement (xls format).

2. The program also extracts cash transactions from BOCHK's cash statement, but does not output it yet.
