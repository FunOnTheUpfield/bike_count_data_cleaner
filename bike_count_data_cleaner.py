#!/usr/bin/env python
# Bike Count Data Cleaner
# Script takes data from a excel spreadsheets and converts it to 'Clean Data' - one observation per row .csv file format.
# By: Simon Stainsby
# Github Username: FunOnTheUpfield
# Created: 13 March 2016
# Last updated: 14 May 2016
# -----------------------------

import json
from xlrd import open_workbook,xldate_as_tuple
import re
from datetime import date,datetime,time


# Initialise variables
EMPTY_VALUE = "NA"
site_id 	= 0
count_id 	= 0


# Configuration -  Data Sources
excelfiles 	= json.load(open('spreadsheetdetails.json'))
outputcols	= json.load(open('outputcolumns.json'))

# function to print output files
def create_output(d,output_file,OUTPUTCOLUMNS):          
     dicvalues = []
     if d != {}:
		for col in sorted(OUTPUTCOLUMNS):
			if OUTPUTCOLUMNS[col] not in d:
				dicvalues.append(EMPTY_VALUE)
			else:
				dicvalues.append(str(d[OUTPUTCOLUMNS[col]]))
		output_file.write("|".join(dicvalues))
		output_file.write("\n") 


# Open a text files for output
siteout  = open('../../Work/output/site_description.txt', 'w') 
siteout.write('site_id | old_id | easting | northing | dist_from_cbd | melway_ref | site_description | suburb |  primary_road | secondary_road')
siteout.write('\n')

countout = open('../../Work/output/count_details.txt', 'w') 
countout.write(' count_id | site_id | survey_year | survey_date | counting | bin_duration | gender_split | male_total | female_total | count_total')
countout.write('\n')

countsummaryout = open('../../Work/output/count_summary.txt', 'w')
countsummaryout.write(' count_id | site_id | survey_date | counting  | count_total | from_north | from_east | from_south | from_west | to_north | to_east | to_south | to_west | male_total | male_from_north | male_from_east | male_from_south | male_from_west | male_to_north | male_to_east | male_to_south | male_to_west | female_total | female_from_north | female_from_east | female_from_south | female_from_west | female_to_north | female_to_east | female_to_south | female_to_west')
countsummaryout.write('\n')


moveout = open('../../Work/output/bike_movement_observations.txt', 'w')
moveout.write('count_id | site_id | bin_start | bin_duration | gender | north_to_west | north_to_south | north_to_east | east_to_north | east_to_west | east_to_south | south_to_east | south_to_north | south_to_west | west_to_south | west_to_east | west_to_north')
moveout.write('\n')


# Iterate through excel data sources scraping cells for interesting values - then summarise.
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
			
			# add results to site details file
			create_output(sitedic,  siteout,  outputcols['SITECOLUMNS'])
			
			# Collect count details
			for k in currentsheet['count_detail_row']:
				count_details 		= currentsheet['count_detail_row'][k]['start_row']
				survey_date_row	= currentsheet['count_details_cells']['survey_date']['row'] + count_details
				survey_date_col 	= currentsheet['count_details_cells']['survey_date']['col']
				survey_date_value 	= sheet.cell(survey_date_row,survey_date_col).value
				
				if survey_date_value != "": 
					# Not every site in the supertue spreadsheet is counted every year
					# The supertue spreadsheet contains blank 'count details forms'.
					# So only attempt to collect movement details if the count date field is populated.
					
					countdic = {}
					countdic['site_id'] = str(site_id)
					count_id = count_id + 1
					countdic['count_id'] = str(count_id)

					countsummarydic = {}
					countsummarydic['site_id'] = str(site_id)
					countsummarydic['count_id'] = str(count_id)

					for k in currentsheet['count_details_cells']:
						row = currentsheet['count_details_cells'][k]['row'] + count_details
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
						countsummarydic['survey_date'] = formatted_date
					except:
						print 'Date problem'
						# TODO: Handle this use case (doesn't come up in the super tue data sheet)
						continue                           
     
					# Some count atributes are implicit and not stored in the super tue spreadsheet
					# Where they are not specified in the supertue spreadsheet, look it up in 'spreadsheetdetails.json'
					for k in currentsheet['count_detail_attributes']:                    
						countdic[k] = currentsheet['count_detail_attributes'][k]['value']
						countsummarydic[k] = currentsheet['count_detail_attributes'][k]['value']
					try:
						countdic['bin_duration'] = int(countdic['bin_duration'])
					except:
						continue


					# Collect the most detailed bike movement information stored in the spreadsheet, individual movements by time bin.

					# Store general information about the count in the first columns of the row
					movedic = {}
					movedic['site_id']  		= str(site_id)
					movedic['count_id'] 		= str(count_id)
					movedic['bin_duration'] 		= countdic['bin_duration'] 

					counttotal = 0

					# Check what information is stored in the supertue spread sheet 
					# Some year's count information might only have 120 minute summaries (ie no 15 min breakdown) and no gender details
					# Other year's count information have 15min breakdowns but no gender breakdown. 

					if countdic['bin_duration'] == 120:
						gender = ('NA',)
						# Skip to year's summary row in the spreadsheet for scrape
						row = count_details + 27
						col = currentsheet['oldcount_summary']['count_total']['col']						
						counttotal = int(sheet.cell(row,col).value)
						for oldie in currentsheet['oldcount_summary']:
							col = currentsheet['oldcount_summary'][oldie]['col']
							countsummarydic[oldie] =	int(sheet.cell(row,col).value)
						
					elif countdic['gender_split'] == 'N':
						gender = ('NA',)
					else:
						gender = ('male','female')
					
					# Collect the most detailed information stored in the messy spreadsheet
					# In super tue worksheet, this information is stored in blocks on the second page of each sheet
					# The file 'spreadsheetdetails.json' specifies where these blocks begin and end.
					
					for mf in gender:
						genderedtotal = 0
						movedic['gender'] = mf
						# TODO :	Fix the hardcoded 7am-9am timerange in the preference file. 
						# 		The Upfield Corridor study includes much longer observation ranges.
						start 	 = currentsheet['movement_bin_row_range']['seven_am_to_nine_am']['start']
						finish	 = currentsheet['movement_bin_row_range']['seven_am_to_nine_am']['finish']
						
						# Collect bin start time
						for r in range(start,finish):         
							row = r + count_details
							col = currentsheet['movement_bin_times']['bin_start']['col']
							movedic['bin_start'] = sheet.cell(row,col).value
							
							# Convert the bin start time from excel into 'YYYY-MM-DD HH:MM:SS' 24hr format
							try:                   
								bin_start_time			= float(movedic['bin_start']) 
								preformatted_start_time	= xldate_as_tuple(bin_start_time,workbook.datemode)
								formatted_time 		= time(*preformatted_start_time[3:5])
								long_start_time		= datetime.combine(formatted_date,formatted_time)
							
								# Use observation start time as an identifier for the row                    
								movedic['bin_start'] 	= long_start_time
 							except:
								continue
							
							# Collect details from each row of observations in the count
							rowtotal = 0
							
							if 	movedic['gender'] != 'female': 
								movement_lookup = 'male_total'
								movements = ('male_total', 
										'male_from_north','male_from_east', 'male_from_south', 'male_from_west',
										'male_to_north', 'male_to_east', 'male_to_south', 'male_to_west')								
								# When no gendersplit information was recorded, total male and female movements 
								# were recorded in the male column.
					
							else:
								movement_lookup = 'female_total'
								movements = ('female_total', 
										'female_from_north','female_from_east', 'female_from_south', 'female_from_west',
										'female_to_north', 'female_to_east', 'female_to_south', 'female_to_west')								


							# Collect most detailed observations availabel
							for i in currentsheet[movement_lookup]:
								col = currentsheet[movement_lookup][i]['col']
								move = sheet.cell(row,col).value
								# Sum rows where possible.
								try:
									rowtotal = rowtotal + int(move)
								except:
									continue

								movedic[i] = move
							
							# Output the row of observations, and the row summaries to text file
							create_output(movedic, moveout, outputcols['MOVECOLUMNS'])

							for summaries in (movements):
								summarytotal = 0
								for i in currentsheet[summaries]:									
									col = currentsheet[summaries][i]['col']
									turn = sheet.cell(row,col).value
									try:
										summarytotal = summarytotal + int(turn)
									except:
										continue
								countsummarydic[summaries] = summarytotal

							#Sum scraped values to create count totals
							genderedtotal = genderedtotal + rowtotal			
						counttotal = counttotal + genderedtotal
						
						if mf == 'male':
							countdic['male_total'] = genderedtotal
						elif mf == 'female':
							countdic['female_total'] = genderedtotal
						else:
							countdic['male_total'] = 'NA'
							countdic['female_total'] = 'NA'

							# Counts with no gender split should have 'NA' rather than '0' in the gender breakdown fields					
							countsummarydic['male_total'] = 'NA'
							countsummarydic['male_from_north'] = 'NA'
							countsummarydic['male_from_east'] = 'NA'
							countsummarydic['male_from_south'] = 'NA'
							countsummarydic['male_from_west'] = 'NA'
							countsummarydic['male_to_north'] = 'NA'
							countsummarydic['male_to_east'] = 'NA'
							countsummarydic['male_to_south'] = 'NA'
							countsummarydic['male_to_west'] = 'NA'
							countsummarydic['female_total'] = 'NA'
							countsummarydic['female_from_north'] = 'NA'
							countsummarydic['female_from_east'] = 'NA'
							countsummarydic['female_from_south'] = 'NA'
							countsummarydic['female_from_west'] = 'NA'
							countsummarydic['female_to_north'] = 'NA'
							countsummarydic['female_to_east'] = 'NA'
							countsummarydic['female_to_south'] = 'NA'
							countsummarydic['female_to_west'] = 'NA'

						# Add the male and female summaries together - where numeric data is available
						try:
							countsummarydic['from_north']	= int(countsummarydic['male_from_north']+ countsummarydic['female_from_north'])
							countsummarydic['from_east'] 	= int(countsummarydic['male_from_east'] + countsummarydic['female_from_east'])
							countsummarydic['from_south'] = int(countsummarydic['male_from_south'] + countsummarydic['female_from_south'])
							countsummarydic['from_west'] 	= int(countsummarydic['male_from_west'] + countsummarydic['female_from_west'])
							countsummarydic['to_north'] 	= int(countsummarydic['male_to_north'] + countsummarydic['female_to_north'])
							countsummarydic['to_east'] 	= int(countsummarydic['male_to_east'] + countsummarydic['female_to_east'])
							countsummarydic['to_south'] 	= int(countsummarydic['male_to_south'] + countsummarydic['female_to_south'])
							countsummarydic['to_west'] 	= int(countsummarydic['male_to_west'] + countsummarydic['female_to_west'])

						except:
							continue
											
					countdic['count_total'] = counttotal
					countsummarydic['count_total'] = counttotal
	

					# add results to count details file
					create_output(countdic,  countout,  outputcols['COUNTCOLUMNS'])
					create_output(countsummarydic, countsummaryout, outputcols['COUNTSUMMARYOUT'])

