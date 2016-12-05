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
Todo
++++++++++
1. The investment lookup function, is repeated with other projects, like trade_converter and jpm, consider move this part to an independant project, so we can centralize lookup and checking.



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
