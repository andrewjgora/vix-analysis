import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import Quandl as q #imports Quandl and assigns 'q' as its shortcut so we can just type q.get() etc
import os
from openpyxl import Workbook
from openpyxl.styles import Style, PatternFill, Color, Border, Side, Alignment


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
	
	colors = ['47B247', #0-4 green, 5 yellow, 6-16 red
			  '52CC52',
			  '5CE65C',
			  '75FF75',
			  '85FF85',
			  'E6E600',
			  '990000',
		  	  'CC0000',
			  'B20000',
		  	  'E60000',
		  	  'FF0000',
		  	  'FF1919',
			  'FF3333',
			  'FF4D4D',
			  'FF6666',
			  'FF8080',
			  'FF9999']

	side = Side(border_style = 'medium', color = 'FF000000')
	xalignment = Alignment(horizontal = 'center')
	xborder = Border(left = side, right = side, top = side, bottom = side)
	styles = [Style(fill = PatternFill(fill_type = 'solid', start_color = colors[0], end_color = colors[0]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[1], end_color = colors[1]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[2], end_color = colors[2]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[3], end_color = colors[3]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[4], end_color = colors[4]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[5], end_color = colors[5]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[6], end_color = colors[6]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[7], end_color = colors[7]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[8], end_color = colors[8]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[9], end_color = colors[9]),
																  border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[10], end_color = colors[10]),
																	border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[11], end_color = colors[11]),
																	border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[12], end_color = colors[12]),
																	border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[13], end_color = colors[13]),
																	border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[14], end_color = colors[14]),
																	border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[15], end_color = colors[15]),
																	border = xborder, alignment = xalignment),
			  Style(fill = PatternFill(fill_type = 'solid', start_color = colors[16], end_color = colors[16]),
																	border = xborder, alignment = xalignment)]
	


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

			#47B247 9 - 10
			#52CC52 between 11-10
			#5CE65C between 12-13
			#75FF75 between 13-14
			#85FF85 between 14-15
			#E6E600 between 15-18		
			xsheet.cell(row = xlrow, column = month+2).value = curx #add value to excel cell
			if curx <= 10:
				xsheet.cell(row = xlrow, column = month+2).style = styles[0]
			elif 10 < curx <= 11:
				xsheet.cell(row = xlrow, column = month+2).style = styles[1]
			elif 11 < curx <= 12:
				xsheet.cell(row = xlrow, column = month+2).style = styles[2]
			elif 12 < curx <= 13:
				xsheet.cell(row = xlrow, column = month+2).style = styles[3]
			elif 13 < curx <= 14:
				xsheet.cell(row = xlrow, column = month+2).style = styles[4]
			elif 14 < curx <= 15:
				xsheet.cell(row = xlrow, column = month+2).style = styles[4]
			elif 15 < curx <= 18:
				xsheet.cell(row = xlrow, column = month+2).style = styles[5]
			elif 18 < curx <= 19:
				xsheet.cell(row = xlrow, column = month+2).style = styles[6]
			elif 19 < curx <= 20:
				xsheet.cell(row = xlrow, column = month+2).style = styles[7]
			elif 20 < curx <= 21:
				xsheet.cell(row = xlrow, column = month+2).style = styles[8]
			elif 21 < curx <= 22:
				xsheet.cell(row = xlrow, column = month+2).style = styles[9]
			elif 22 < curx <= 23:
				xsheet.cell(row = xlrow, column = month+2).style = styles[10]
			elif 23 < curx <= 24:
				xsheet.cell(row = xlrow, column = month+2).style = styles[11]
			elif 24 < curx <= 25:
				xsheet.cell(row = xlrow, column = month+2).style = styles[12]
			elif 25 < curx <= 26:
				xsheet.cell(row = xlrow, column = month+2).style = styles[13]
			elif 26 < curx <= 27:
				xsheet.cell(row = xlrow, column = month+2).style = styles[14]
			elif 27 < curx <= 28:
				xsheet.cell(row = xlrow, column = month+2).style = styles[15]
			elif curx > 28:
				xsheet.cell(row = xlrow, column = month+2).style = styles[16]

	xlbook.save('%s.xlsx' % sheetname)

printexcelsheet(1990, 2015, 'All_Lows')
