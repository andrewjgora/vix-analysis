import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import Quandl as q #imports Quandl and assigns 'q' as its shortcut so we can just type q.get() etc
import os

vix = q.get("YAHOO/INDEX_VIX")
#vixmonthly = q.get("YAHOO/INDEX_VIX", collapse="monthly")
#allhilows = vixdata.loc[:,['High', 'Low']]

#path = os.path.expanduser('~/vix-analysis/monthly.xlsx')
#vix2 = pd.read_excel('monthly.xlsx')

startyear = 1990
endyear = 2015
#leap years: 1992, 1996, 2000, 2004, 2008, 2012

monthlim = {1: 31 , 2: 28, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}
tomonth = {1: 'January',
		   2: 'February',
		   3: 'March',
		   4: 'April',
		   5: 'May',
		   6: 'June',
		   7: 'July',
		   8: 'August',
		   9: 'September',
		   10: 'October',
		   11: 'November',
		   12: 'December'}

def printextrema():
	lowdate = ''
	highdate = ''

	for curyear in range(startyear, endyear + 1):
		for curmonth in range(0, 12):
			for curday in range(0, monthlim[curmonth + 1]): #find lowest low & highest high of each month
				try:
					lowval = 999
					highval = 0
					if curday == 0:
						lowval = vix.loc['%d-%d-%d' % (curyear, curmonth + 1, curday + 1),'Low']
						highval = vix.loc['%d-%d-%d' % (curyear, curmonth + 1, curday + 1),'High']
					else:
						nextlow = vix.loc['%d-%d-%d' % (curyear, curmonth + 1, curday + 1),'Low']
						nexthigh = vix.loc['%d-%d-%d' % (curyear, curmonth + 1, curday + 1),'High']				
						if nextlow < lowval:
							lowdate = '%d-%d-%d' % (curyear, curmonth + 1, curday + 1)
							lowval = nextlow
						if nexthigh > highval:
							highdate = '%d-%d-%d' % (curyear, curmonth + 1, curday + 1)
							highval = nexthigh
				except KeyError:
					hi = 1
			print('Lowest Low for %s: %d' % (lowdate, lowval))
			print('Highest High for %s: %d' % (highdate, highval))
			print()
#s.to_frame(name='column_name').to_excel('xlfile.xlsx', sheet_name='s')

printextrema()