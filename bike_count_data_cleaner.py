#!/usr/bin/env python
# Bike Count Data Cleaner
# Script takes data from a excel spreadsheets and converts it to 'Clean Data' - one observation per row .csv file format.
# By: Simon Stainsby
# Github Username: FunOnTheUpfield
# Created: 13 March 2016
# Last updated: 13 March 2016
# -----------------------------

import json
from xlrd import open_workbook,xldate_as_tuple
import re
from datetime import date,datetime,time

EMPTY_VALUE = "NA"

SITECOLUMNS = [
	'site_id', 
	'old_id',          
	'easting',
	'northing',
	'dist_from_cbd', 
	'melway_ref', 
	'site_description',
	'suburb', 
	'primary_road', 
	'secondary_road',

	]

# Open a text files for output
siteout  = open('../../Work/output/site_description.txt', 'w') 
siteout.write('site_id | old_id | easting | northing | dist_from_cbd | melway_ref | site_description | suburb |  primary_road | secondary_road')
siteout.write('\n')

# function to print output files
def create_output(d,output_file,OUTPUTCOLUMNS):          
     dicvalues = []
     if d != {}:
          for col in OUTPUTCOLUMNS:
               if col not in d:
                    dicvalues.append(EMPTY_VALUE)
               else:
                    dicvalues.append(str(d[col]))
          output_file.write("|".join(dicvalues))
          output_file.write("\n") 



# Scraper 
excelfiles = json.load(open('spreadsheetdetails.json'))

site_id = 0
allsites = {}

for spreadsheet in excelfiles:
	if spreadsheet != "metadata":
		sourcefile = excelfiles[spreadsheet]["filepath"] + excelfiles[spreadsheet]["filename"]
		workbook	 = open_workbook(sourcefile, on_demand=True)
		print sourcefile, ' Open'

		currentsheet = excelfiles[spreadsheet]
		start 	 = currentsheet['worksheet_range']['start']
		finish	 = currentsheet['worksheet_range']['finish']
	
		for worksheet_num in range(start, finish):
			sheet = workbook.sheet_by_index(worksheet_num)

			if spreadsheet == "supertue":
				# For the super tuesday sheet, collect site details from top of worksheet
				sitedic  = {}
				for k in currentsheet['site_detail_cell']:
					row = currentsheet['site_detail_cell'][k]['row']
					col = currentsheet['site_detail_cell'][k]['col']
					if row is None:
						sitedic[k] = EMPTY_VALUE
					else:
						sitedic[k] = sheet.cell(row,col).value

			
				site_id = site_id +1
				# Assumes all counts for a given site are on a single worksheet
				# This is true for super tue 7am-9am count but not for the upfield corridor count

			sitedic['site_id']   = str(site_id)

			allsites[str(site_id)] = sitedic

			# add results to site details file
			create_output(sitedic,  siteout,  SITECOLUMNS)



