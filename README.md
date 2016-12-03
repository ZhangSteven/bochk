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
1. The program now output ISIN and Bloomberg FIGI as the identifier. It's OK for the concord fund. But for those HTM funds, because a HTM bond will not have ISIN in the security master if it is also in a AFS position in another fund. So we need to output another "geneva_investment_id" column, we can get rid of the "accounting treatment" column.


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
