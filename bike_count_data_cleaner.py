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

# TODO: Move these constancts to an conf or ini file
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
	'secondary_road'
	]

COUNTCOLUMNS = [
	'count_id',
	'site_id',
	'survey_year',
	'survey_date',
	'counting',
	'bin_duration',
	'gender_split'
	]

# Open a text files for output
siteout  = open('../../Work/output/site_description.txt', 'w') 
siteout.write('site_id | old_id | easting | northing | dist_from_cbd | melway_ref | site_description | suburb |  primary_road | secondary_road')
siteout.write('\n')

countout = open('../../Work/output/count_details.txt', 'w') 
countout.write(' count_id | site_id | survey_year | survey_date | counting | bin_duration | gender_split')
countout.write('\n')


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

site_id 	= 0
count_id 	= 0
allsites 	= {}
allcounts	= {}

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
					row 	= currentsheet['site_detail_cell'][k]['row']
					col 	= currentsheet['site_detail_cell'][k]['col']
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

			## Collect count details
			## Scrape the six places where count details might be stored, but only update the dictionary if there are values there.
			## Not every site is counted every year, and some sites are counted more often than others.
			
			for k in currentsheet['count_detail_row']:
				count_details 		= currentsheet['count_detail_row'][k]['start_row']
				survey_date_row	= currentsheet['count_details_cells']['survey_date']['row'] + count_details
				survey_date_col 	= currentsheet['count_details_cells']['survey_date']['col']
				survey_date_value 	= sheet.cell(survey_date_row,survey_date_col).value

				if survey_date_value != "": 
					countdic = {}
					countdic['site_id'] = str(site_id)
					count_id = count_id + 1
					countdic['count_id'] = str(count_id)
					for k in currentsheet['count_details_cells']:
						row = currentsheet['count_details_cells'][k]['row']+ count_details
						col = currentsheet['count_details_cells'][k]['col']
						countdic[k] = sheet.cell(row,col).value 

     
					# Convert count date from excel format to YYYY-MM-DD
					try:                    
						count_date = float(countdic['survey_date'])
						preformatted_date = xldate_as_tuple(count_date,workbook.datemode)
						formatted_date = date(*preformatted_date[0:3])
						survey_year = preformatted_date[0]                                                 
						countdic['survey_date'] = formatted_date
						countdic['survey_year'] = survey_year
					except:
						print 'Date problem'
						continue                           
                              
					# Collect count attributes stored in preferences
					for k in currentsheet['count_detail_attributes']:                    
						countdic[k] = currentsheet['count_detail_attributes'][k]['value']
					try:
						countdic['bin_duration'] = int(countdic['bin_duration'])
					except:
						continue

					allcounts[str(count_id)] = countdic
					# add results to count details file
					create_output(countdic,  countout,  COUNTCOLUMNS)

print allcounts
