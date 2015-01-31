import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import Quandl as q #imports Quandl and assigns 'q' as its shortcut so we can just type q.get() etc
import os
from openpyxl import Workbook

vix = q.get("YAHOO/INDEX_VIX")
#vixmonthly = q.get("YAHOO/INDEX_VIX", collapse="monthly")
#allhilows = vixdata.loc[:,['High', 'Low']]

#path = os.path.expanduser('~/vix-analysis/monthly.xlsx')
#vix2 = pd.read_excel('monthly.xlsx')


#leap years: 1992, 1996, 2000, 2004, 2008, 2012

monthlim = {1: 31 , 2: 28, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}
xltomonth = {2: 'January',
		  	    3: 'February',
		   		4: 'March',
		   		5: 'April',
		   		6: 'May',
		   		7: 'June',
		   		8: 'July',
		   		9: 'August',
		   		10: 'September',
		   		11: 'October',
		   		12: 'November',
		   		13: 'December'}

def printextrema(startyear, endyear):

	curlow = np.nan
	curhigh = np.nan

	for year in range(startyear, endyear):
		#add new row to lows spreadsheet
		#add new row to highs spreadsheet
		for month in range(12):
			for day in range(monthlim[month + 1]):
				if year == startyear:
					day += 1
					curlow = 999
					curhigh = 0
				curdatestr = '%d-%d-%d' % (year, month + 1, day + 1)
				lastopen = vix.index.asof(curdatestr)
				curdate = pd.to_datetime(curdatestr)

				if curdate.is_month_start:
					curlow = 999
					curhigh = 0
				nextlow = vix.loc[lastopen, 'Low']
				nexthigh = vix.loc[lastopen, 'High']
				if nextlow < curlow:
					lowdate = curdate
					curlow = nextlow
				if nexthigh > curhigh:
					highdate = curdate
					curhigh = nexthigh
				if curdate.is_month_end:
					break
			print('_______________________________________')
			print()
			print('Extrema for %s, %d:' % (xltomonth[month + 1], year))
			print()
			print('Lowest Low: $%d on %s %d' % (curlow, xltomonth[lowdate.month], lowdate.day))
			print('Highest High: $%d on %s %d' %  (curhigh, xltomonth[highdate.month], highdate.day))
			print('_______________________________________')
			#
	


def printexcelsheet(startyear, endyear, sheetname):

	curlow = np.nan
	xlbook = Workbook()
	xsheet = xlbook.active
	xsheet.title = sheetname
	# compare = {'High': }

	# def high(val):
	# 	if val > curx:
	# 		xdate = curdate
	# 		curx =

	# create the cells for January-December
	for i in range(2, 14):
		xsheet.cell(row = 1, column = i).value = xltomonth[i]

	for year in range(startyear, endyear):

		xlrow = year - startyear + 2
		xsheet.cell(row = xlrow, column = 1).value = year
		for month in range(12):
			for day in range(monthlim[month + 1]):
				if year == startyear:
					day += 1
					curx = 999
				curdatestr = '%d-%d-%d' % (year, month + 1, day + 1)
				lastopen = vix.index.asof(curdatestr)
				curdate = pd.to_datetime(curdatestr)
				if curdate.is_month_start:
					curx = 999
				nextx = vix.loc[lastopen, 'Low']
				if nextx < curx:
					xdate = curdate
					curx = nextx
				if curdate.is_month_end: #we know we found the lowest now, so break
					break
			xsheet.cell(row = xlrow, column = month+2).value = curx #add value to excel cell

	xlbook.save('%s.xlsx' % sheetname)

printexcelsheet(1990, 2015, 'All_Lows')
