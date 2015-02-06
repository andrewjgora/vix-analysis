import numpy as np
#import matplotlib.pyplot as plt
import pandas as pd
import Quandl as q #imports Quandl and assigns 'q' as its shortcut so we can just type q.get() etc
from openpyxl import Workbook
from openpyxl.styles import Style, PatternFill, Color, Border, Side, Alignment


vix = q.get("YAHOO/INDEX_VIX")


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
	


def printexcelsheet(startyear, endyear, sheetname, start, hold, buy, end):

	xlbook = Workbook()
	xsheet = xlbook.active
	xsheet.title = sheetname
	
	# green 00FF00-70FF00 = 112 values/8 = 14 shades
 	# yellow DFFF00-FFDF00 = 64 values/8 = 8 shades
	# red 	FF6000-FF0000 = 96 values/8 = 12 shades

	colors = []

	for i in range(0, 112, 8):
		if i <= 8:
			colors.append('0%XFF00' % i)
		else:
			colors.append('%XFF00' % i)

	for j in range(223, 255, 8):
		colors.append('%XFF00' % j)
	for j in reversed(range(223, 255, 8)):
		colors.append('FF%X00' % j)

	for k in reversed(range(0, 96, 8)):
		if k <= 8:
			colors.append('FF0%X00' % k)
		else:
			colors.append('FF%X00' % k)

	greenlist = np.linspace(start, hold, 14) # numpy linear space
	yellowlist = np.linspace(hold, buy, 8)	 # linspace(a,b,c) returns c evenly distributed values
	redlist = np.linspace(buy, end, 12)		 # from a to b, including endpoints a and b

	# concatenate the three lists and call unique() to get rid of duplicates
	comparisonlist = np.unique(np.concatenate((greenlist, yellowlist, redlist))) #size = 34
	length = len(comparisonlist)

	side = Side(border_style = 'medium', color = 'FF000000')
	xalignment = Alignment(horizontal = 'center')
	xborder = Border(left = side, right = side, top = side, bottom = side)

	#function to create a style from the passed cell color hex
	def coloredcell(color): 
		return Style(fill = PatternFill(fill_type = 'solid', start_color = color, end_color = color), border = xborder, alignment = xalignment)

	# make a style for every element of colors using coloredcell
	styles = [coloredcell(color) for color in colors] 

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
		
			xsheet.cell(row = xlrow, column = month+2).value = curx #add value to cell

			#now color
			# find the index 'z' of the first value 'ceiling' of comparisonlist that is greater than curx
			# set cell style to corresponding colored cell style from styles
			for ceiling in comparisonlist:
				if curx < ceiling: 
					z = np.where(comparisonlist==ceiling)
					xsheet.cell(row = xlrow, column = month+2).style = styles[z[0]]
					break
				elif curx > comparisonlist[length - 1]: # we need to check if curx is greater than upper bound too
					xsheet.cell(row = xlrow, column = month+2).style = styles[length - 1]
					break

	xlbook.save('%s.xlsx' % sheetname)

printexcelsheet(1990, 2015, 'All_Lows', 10, 15, 18, 28)
