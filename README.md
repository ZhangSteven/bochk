# bochk

This is to convert the BOCHK broker statement into files containing investment positions and cash to reconcile with Advent Geneva system.

The positions are trade day positions, cash is settlement day cash.


++++++++++
Note
++++++++++
1. The program hard coded some of the holding account name and cash name, to be mapped to portfolio ids. The following portfolios/account names not mapped yet:

	12528's holding account name not known
	16454 not mapped
	Greenblue fund not mapped.

	see map_holding_to_portfolio_id() and map_cash_to_portfolio_id().

2. The program assumes a portfolio holds either all HTM or all Trading bonds, see lookup_portfolio_accounting_treatment().
