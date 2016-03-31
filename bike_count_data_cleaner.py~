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

MOVECOLUMNS = [
          'count_id',
          'site_id',
          'bin_start',
		'bin_duration',
          'male_north_to_south',
          'male_north_to_east',
          'male_east_to_north',
          'male_east_to_west',
          'male_east_to_south',
          'male_south_to_east',
          'male_south_to_north',
          'male_south_to_west',
          'male_west_to_south',
          'male_west_to_east',
          'male_west_to_north',
          'male_north_to_west',
          'female_north_to_south',
          'female_north_to_east',
          'female_east_to_north',
          'female_east_to_west',
          'female_east_to_south',
          'female_south_to_east',
          'female_south_to_north',
          'female_south_to_west',
          'female_west_to_south',
          'female_west_to_east',
          'female_west_to_north',
          'female_north_to_west'
          ]

BINTOTALS = [
          'count_id',
          'site_id',
          'bin_start',
		'bin_duration',
			# All movements
		#'all_movements_at_intersection',
			# By_gender
		#'total_male_using_intersection',
		#'total_female_using_intersection',
			# All From north
		'all_from_north_to_west',
		'all_from_north_to_south',
		#'all_from_north_to_east',
			# All From east
		#'all_from_east_to_north',
		#'all_from_east_to_west',
		#'all_from_east_to_south',
			# All From south
		#'all_from_south_to_east',
		#'all_from_south_to_north',
		#'all_from_south_to_west',
		# All From west,
		#'all_from_west_to_south',
		#'all_from_west_to_east',
		#'all_from_west_to_north',
			# All From
		#'all_from_east',
		#'all_from_south',
		#'all_from_west',
			# All Heading to
		#'all_to_north',
		#'all_to_east',
		#'all_to_south',
		#'all_to_west',
			# Total Traffic volume on road segment
		#'all_on_northern_road',
		#'all_on_eastern_road',
		#'all_on_southern_road',
		#'all_on_western_road',
			# Female From north
		#'female_from_north_to_west',
		#'female_from_north_to_south',
		#'female_from_north_to_east',
			# Female From east
		#'female_from_east_to_north',
		#'female_from_east_to_west',
		#'female_from_east_to_south',
			# Female From south
		#'female_from_south_to_east',
		#'female_from_south_to_north',
		#'female_from_south_to_west',
			# Female From west,
		#'female_from_west_to_south',
		#'female_from_west_to_east',
		#'female_from_west_to_north',
			# Female From
		#'female_from_east',
		#'female_from_south',
		#'female_from_west',
			# Female Heading to
		#'female_to_north',
		#'female_to_east',
		#'female_to_south',
		#'female_to_west',
			# Female Traffic volume on road segment
		#'female_on_northern_road',
		#'female_on_eastern_road',
		#'female_on_southern_road',
		#'female_on_western_road'
		]

# Open a text files for output
siteout  = open('../../Work/output/site_description.txt', 'w') 
siteout.write('site_id | old_id | easting | northing | dist_from_cbd | melway_ref | site_description | suburb |  primary_road | secondary_road')
siteout.write('\n')

countout = open('../../Work/output/count_details.txt', 'w') 
countout.write(' count_id | site_id | survey_year | survey_date | counting | bin_duration | gender_split')
countout.write('\n')

moveout = open('../../Work/output/bike_movement_observations.txt', 'w')
moveout.write(' count_id | site_id | bin_start | bin_duration | male_north_to_south | male_north_to_east | male_east_to_north | male_east_to_west | male_east_to_south | male_south_to_east | male_south_to_north | male_south_to_west | male_west_to_south | male_west_to_east | male_west_to_north | male_north_to_west | female_north_to_south | female_north_to_east | female_east_to_north | female_east_to_west | female_east_to_south | female_south_to_east |female_south_to_north | female_south_to_west | female_west_to_south | female_west_to_east | female_west_to_north | female_north_to_west')
moveout.write('\n')

bintotalout = open('../../Work/output/bike_movement_summary1.txt', 'w')
bintotalout.write(' count_id | site_id | bin_start | bin_duration | all_from_north_to_west | all_from_north_to_south')
bintotalout.write('\n')

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
excelfiles 	= json.load(open('spreadsheetdetails.json'))

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
					except:
						print 'Date problem'
						# TODO: do something ???
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


					# Collect move details
					movedic = {}
					movedic['site_id']  = countdic['site_id']
					movedic['count_id'] = countdic['count_id']
					movedic['bin_duration'] = countdic['bin_duration'] 
					seven_to_nine_moves = {}

					# Check if all all detailed count information available.
					if countdic['bin_duration'] == 120 and countdic['gender_split'] == 'Y':
						print 'Old super tuesday count, go to summary row'						
						# TODO: Actually do something here
					else:
						start 	 = currentsheet['movement_bin_row_range']['seven_am_to_nine_am']['start']
						finish	 = currentsheet['movement_bin_row_range']['seven_am_to_nine_am']['finish']
						# Site has detailed '15min bin' bike movement information.						
						# TODO Fix hardcoded timerange - no good for upfield study.

						for k in range(start,finish):         
								row = k + count_details
								col = currentsheet['movement_bin_times']['bin_start']['col']
								movedic['bin_start'] = sheet.cell(row,col).value                                        
								for i in currentsheet['movement_bin_columns']:
									col = currentsheet['movement_bin_columns'][i]['col']
									move = sheet.cell(row,col).value
									try:
										move = int(move)
									except:
										continue
									movedic[i] = move

									# Convert the bin start time from excel into HH:MM 24hr format
									try:                    
										bin_start_time = float(movedic['bin_start']) 
										preformatted_start_time = xldate_as_tuple(bin_start_time,workbook.datemode)
										formatted_time = time(*preformatted_start_time[3:5])
										long_start_time = datetime.combine(formatted_date,formatted_time)                    
										movedic['bin_start'] = long_start_time
										seven_to_nine_moves['binstart'] = movedic
									except:
										continue

									# add results to count details file
									create_output(movedic, moveout, MOVECOLUMNS)
								
								# Add 'move' values together to create useful summary information (for 15 min count bins)
								bin_totals = {}
								bin_totals['site_id'] 		= countdic['site_id']
								bin_totals['count_id'] 		= countdic['count_id']
								bin_totals['bin_start']		= movedic['bin_start']
								bin_totals['bin_duration']	= movedic['bin_duration']

								summaryfields = 	{
									'all_from_north_to_west' : ('male_north_to_west','female_north_to_west'),
									'all_from_north_to_south': ('male_north_to_south','female_north_to_south')
												}
								

								for summarise in summaryfields:
									fieldsum = 0
									for i in range(0,len(summaryfields[summarise])):
										interestingfield = summaryfields[summarise][i]
										try:
											fieldsum = fieldsum + movedic[interestingfield]
										except:
											continue
									print summarise, fieldsum

									bin_totals[summarise] = fieldsum  


								
 


								#print bin_totals
								create_output(bin_totals,  bintotalout, BINTOTALS)

