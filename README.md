# bochk

This is to convert the BOCHK broker statement into files containing investment positions and cash to reconcile with Advent Geneva system.

The positions are trade day positions, cash is settlement day cash.


++++++++++
Note
++++++++++
1. The program hard coded some of the holding account name and cash name, to be mapped to portfolio ids, look out when new portfolios are added.

2. The program assumes a portfolio holds either all HTM or all Trading bonds, see lookup_accounting_treatment().

3. When a position's security id type is not ISIN, the program then looks up the isin code or BBG ID in investmentLookup.xls file. Currently there are 2 such bonds (based on 2016-7-6 broker statement), when more such bonds appear, update this file.


++++++++++
ver 0.1
++++++++++
1. Generates a cash and holdings csv file from BOCHK's holdings and cash statement (xls format).

2. The program also extracts cash transactions from BOCHK's cash statement, but does not output it yet.
