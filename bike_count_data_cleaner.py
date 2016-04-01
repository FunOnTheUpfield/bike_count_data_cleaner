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


# Initialise variables
EMPTY_VALUE = "NA"
site_id 	= 0
count_id 	= 0
allsites 	= {}
allcounts	= {}


# Configuration -  Data Sources
excelfiles 	= json.load(open('spreadsheetdetails.json'))

# Configuration - Observation summaries
#TODO Move this to an external configuration file
summary_calculations = {
	# All movements
	"all_movements_at_intersection"	: 	(
									"male_north_to_west" , "female_north_to_west"
									"male_north_to_south", "female_north_to_south",
									"male_north_to_east" , "female_north_to_east",
									"male_east_to_north" , "female_east_to_north",
									"male_east_to_west"  , "female_east_to_west",
									"male_east_to_south" , "female_east_to_south",
									"male_south_to_east" , "female_south_to_east",
									"male_south_to_north", "female_south_to_north",
									"male_south_to_west" , "female_south_to_west",
									"male_west_to_south" , "female_west_to_south",
									"male_west_to_east"  , "female_west_to_east",
									"male_west_to_north" , "female_west_to_north"
									),
	# By_gender		
	"total_female_using_intersection"	: 	(
									"female_north_to_west" , "female_north_to_south", "female_north_to_east" ,
									"female_east_to_north" , "female_east_to_west"  , "female_east_to_south" , 
									"female_south_to_east" , "female_south_to_north", "female_south_to_west" , 
									"female_west_to_south" , "female_west_to_east"  , "female_west_to_north" ,  
									),

	"total_male_using_intersection"	: 	(
									"male_north_to_west" , "male_north_to_south", "male_north_to_east" ,
									"male_east_to_north" , "male_east_to_west"  , "male_east_to_south" , 
									"male_south_to_east" , "male_south_to_north", "male_south_to_west" , 
									"male_west_to_south" , "male_west_to_east"  , "male_west_to_north" ,  
									),
	# Use by road segment
	"all_on_northern_road"			: 	("male_north_to_west" , "female_north_to_west",
									"male_north_to_south", "female_north_to_south"
									"male_north_to_east" , "female_north_to_east",
									"male_east_to_north" , "female_east_to_north",
									"male_south_to_north", "female_south_to_north",
									"male_west_to_north" , "female_west_to_north"
									),

	"all_on_eastern_road"			: 	(
									"male_east_to_north" , "female_east_to_north", 
									"male_east_to_west" , "female_east_to_west", 
									"male_east_to_south" , "female_east_to_south",
									"male_north_to_east" , "female_north_to_east",
									"male_south_to_east" , "female_south_to_east",
									"male_west_to_east"  , "female_west_to_east"
									),
	
	"all_on_southern_road"			: 	(
									"male_south_to_east" , "female_south_to_east",
									"male_south_to_north", "female_south_to_north",
									"male_south_to_west" , "female_south_to_west",
									"male_north_to_south", "female_north_to_south",
									"male_east_to_south" , "female_east_to_south",
									"male_west_to_south" , "female_west_to_south"
									),
	
	"all_on_western_road"			: 	(
									"male_west_to_south" , "female_west_to_south",
									"male_west_to_east"  , "female_west_to_east",
									"male_west_to_north" , "female_west_to_north",
									"male_north_to_west" , "female_north_to_west",
									"male_east_to_west"  , "female_east_to_west",
									"male_south_to_west" , "female_south_to_west"
									),
	# Origin of traveller
	"all_from_north"				: 	(
									"male_north_to_west" , "female_north_to_west",
									"male_north_to_south", "female_north_to_south"
									"male_north_to_east" , "female_north_to_east"
									),
								
	"all_from_east"				: 	(
									"male_east_to_north" , "female_east_to_north", 
									"male_east_to_west" , "female_east_to_west", 
									"male_east_to_south" , "female_east_to_south"
									),
								
	"all_from_south"				: 	(
									"male_south_to_east" , "female_south_to_east",
									"male_south_to_north", "female_south_to_north",
									"male_south_to_west" , "female_south_to_west"
									),
								
	"all_from_west"				: 	(
									"male_west_to_south" , "female_west_to_south",
									"male_west_to_east"  , "female_west_to_east",
									"male_west_to_north" , "female_west_to_north"
									),
	# Destination of traveller
	"all_to_north"					: 	(
									"male_east_to_north" , "female_east_to_north",
									"male_south_to_north", "female_south_to_north",
									"male_west_to_north" , "female_west_to_north"
									),

	"all_to_east"					: 	(
									"male_north_to_east" , "female_north_to_east",
									"male_south_to_east" , "female_south_to_east",
									"male_west_to_east"  , "female_west_to_east"
									),
								
	"all_to_south"					: 	(
									"male_north_to_south", "female_north_to_south",
									"male_east_to_south" , "female_east_to_south",
									"male_west_to_south" , "female_west_to_south"
									),
								
	"all_to_west"					: 	(
									"male_north_to_west" , "female_north_to_west",
									"male_east_to_west"  , "female_east_to_west",
									"male_south_to_west" , "female_south_to_west"
									),
	# Female travellers by road segment
	"female_on_northern_road"		: 	("female_north_to_west", "female_north_to_south","female_north_to_east",
									"female_east_to_north", "female_south_to_north", "female_west_to_north"
									),
		
	"female_on_eastern_road"			: 	("female_east_to_north",  "female_east_to_west", "female_east_to_south",
									"female_north_to_east","female_south_to_east","female_west_to_east"
									),

	"female_on_southern_road"		: 	("female_south_to_east","female_south_to_north", "female_south_to_west",
									"female_north_to_south", "female_east_to_south", "female_west_to_south"
									),

	"female_on_western_road"			: 	("female_west_to_north","female_west_to_east", "female_west_to_south", 
									"female_north_to_west","female_east_to_west", "female_south_to_west"
									),

	# Origin of Female travellers
	"female_from_north"				: 	("female_north_to_west", "female_north_to_south", "female_north_to_east"),
	
	"female_from_east"				: 	("female_east_to_north", "female_east_to_west" , "female_east_to_south"),

	"female_from_south"				: 	("female_south_to_east","female_south_to_north","female_south_to_west"), 
									
	"female_from_west"				: 	("female_west_to_south", "female_west_to_east", "female_west_to_north"),

	# Destination of female travellers
									
	"female_to_north"				: 	("female_east_to_north","female_south_to_north", "female_west_to_north"),
									
	"female_to_east"				: 	("female_north_to_east","female_south_to_east", "female_west_to_east"),

	"female_to_south"				: 	("female_north_to_south", "female_east_to_south", "female_west_to_south"),
									
	"female_to_west"				: 	("female_north_to_west", "female_east_to_west", "female_south_to_west"),

	# Male travellers by road segment
	"male_on_northern_road"			: 	("male_north_to_west", "male_north_to_south","male_north_to_east",
									"male_east_to_north", "male_south_to_north", "male_west_to_north"
									),
		
	"male_on_eastern_road"			: 	("male_east_to_north",  "male_east_to_west", "male_east_to_south",
									"male_north_to_east","male_south_to_east","male_west_to_east"	
									),

	"male_on_southern_road"			: 	("male_south_to_east","male_south_to_north", "male_south_to_west",
									"male_north_to_south", "male_east_to_south", "male_west_to_south"
									),

	"male_on_western_road"			: 	("male_west_to_north","male_west_to_east", "male_west_to_south", 
									"male_north_to_west","male_east_to_west", "male_south_to_west"
									),
	# Origin of male travellers
	"male_from_north"			 	: 	("male_north_to_west", "male_north_to_south", "male_north_to_east"),

	"male_from_east"				: 	("male_east_to_north", "male_east_to_west" , "male_east_to_south"),

	"male_from_south"				: 	("male_south_to_east","male_south_to_north","male_south_to_west"), 
									
	"male_from_west"				: 	("male_west_to_south", "male_west_to_east", "male_west_to_north"),
									
	"male_to_north"				: 	("male_east_to_north","male_south_to_north", "male_west_to_north"),
									
	# Destination of male travellers
	"male_to_east"					: 	("male_north_to_east","male_south_to_east", "male_west_to_east"),

	"male_to_south"				: 	("male_north_to_south", "male_east_to_south", "male_west_to_south"),
									
	"male_to_west"					: 	("male_north_to_west", "male_east_to_west", "male_south_to_west"),

	# Turning movements - without gender split
	"all_from_north_to_west"			: 	("male_north_to_west" , "female_north_to_west"),
	"all_from_north_to_south"		: 	("male_north_to_south", "female_north_to_south"),
	"all_from_north_to_east"			: 	("male_north_to_east" , "female_north_to_east"),
	
	"all_from_east_to_north"			: 	("male_east_to_north" , "female_east_to_north"),
	"all_from_east_to_west"			: 	("male_east_to_west"  , "female_east_to_west"),
	"all_from_east_to_south"			: 	("male_east_to_south" , "female_east_to_south"),
		
	"all_from_south_to_east"			: 	("male_south_to_east" , "female_south_to_east"),
	"all_from_south_to_north"		: 	("male_south_to_north", "female_south_to_north"),
	"all_from_south_to_west"			: 	("male_south_to_west" , "female_south_to_west"),
		
	"all_from_west_to_south"			: 	("male_west_to_south" , "female_west_to_south"),
	"all_from_west_to_east"			: 	("male_west_to_east"  , "female_west_to_east"),
	"all_from_west_to_north"			: 	("male_west_to_north" , "female_west_to_north"),
	
}






# Configuration - Output text file columns 
# TODO: Move these constants to an conf or ini file
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
		'all_movements_at_intersection',
			# By_gender
		'total_male_using_intersection',
		'total_female_using_intersection',
			# Use by road segment
		'all_on_northern_road',
		'all_on_eastern_road',
		'all_on_southern_road',
		'all_on_western_road',
			# Origin of traveller 
		'all_from_north',
		'all_from_east',
		'all_from_south',
		'all_from_west',
			# Destination of traveller
		'all_to_north',
		'all_to_east',
		'all_to_south',
		'all_to_west',
			# Female travellers by road segment
		'female_on_northern_road',
		'female_on_eastern_road',
		'female_on_southern_road',
		'female_on_western_road',
			# Origin of female travellers
		'female_from_north',
		'female_from_east',
		'female_from_south',
		'female_from_west',
			# Destination of female travellers
		'female_to_north',
		'female_to_east',
		'female_to_south',
		'female_to_west',
			# Male travellers by road segment
		'male_on_northern_road',
		'male_on_eastern_road',
		'male_on_southern_road',
		'male_on_western_road',
			# Origin of male travellers
		'male_from_north',
		'male_from_east',
		'male_from_south',
		'male_from_west',
			# Destination of male travellers
		'male_to_north',
		'male_to_east',
		'male_to_south',
		'male_to_west',
			# Turning movements - without gender split
		'all_from_north_to_west',
		'all_from_north_to_south',
		'all_from_north_to_east',
		'all_from_east_to_north',
		'all_from_east_to_west',
		'all_from_east_to_south',
		'all_from_south_to_east',
		'all_from_south_to_north',
		'all_from_south_to_west',
		'all_from_west_to_south',
		'all_from_west_to_east',
		'all_from_west_to_north',
		]

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


# Open a text files for output
siteout  = open('../../Work/output/site_description.txt', 'w') 
siteout.write('site_id | old_id | easting | northing | dist_from_cbd | melway_ref | site_description | suburb |  primary_road | secondary_road')
siteout.write('\n')

countout = open('../../Work/output/count_details.txt', 'w') 
countout.write(' count_id | site_id | survey_year | survey_date | counting | bin_duration | gender_split')
countout.write('\n')

moveout = open('../../Work/output/bike_movement_observations.txt', 'w')
moveout.write('count_id | site_id | bin_start | bin_duration | male_north_to_south | male_north_to_east | male_east_to_north | male_east_to_west | male_east_to_south | male_south_to_east | male_south_to_north | male_south_to_west | male_west_to_south | male_west_to_east | male_west_to_north | male_north_to_west | female_north_to_south | female_north_to_east | female_east_to_north | female_east_to_west | female_east_to_south | female_south_to_east |female_south_to_north | female_south_to_west | female_west_to_south | female_west_to_east | female_west_to_north | female_north_to_west')
moveout.write('\n')

bintotalout = open('../../Work/output/bike_movement_summary_15min.txt', 'w')
bintotalout.write('count_id | site_id | bin_start | bin_duration | all_movements_at_intersection | total_male_using_intersection | total_female_using_intersection | all_on_northern_road | all_on_eastern_road | all_on_southern_road | all_on_western_road | all_from_north | all_from_east | all_from_south | all_from_west | all_to_north | all_to_east | all_to_south | all_to_west | female_on_northern_road | female_on_eastern_road | female_on_southern_road | female_on_western_road | female_from_north | female_from_east | female_from_south | female_from_west | female_to_north | female_to_east | female_to_south | female_to_west | male_on_northern_road | male_on_eastern_road | male_on_southern_road | male_on_western_road | male_from_north | male_from_east | male_from_south | male_from_west | male_to_north | male_to_east | male_to_south | male_to_west | all_from_north_to_west | all_from_north_to_south | all_from_north_to_east | all_from_east_to_north | all_from_east_to_west | all_from_east_to_south | all_from_south_to_east | all_from_south_to_north | all_from_south_to_west | all_from_west_to_south | all_from_west_to_east | all_from_west_to_north')
bintotalout.write('\n')

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
     
					# Collect attributes for each count stored in 'spreadsheetdetails.json'
					for k in currentsheet['count_detail_attributes']:                    
						countdic[k] = currentsheet['count_detail_attributes'][k]['value']
					try:
						countdic['bin_duration'] = int(countdic['bin_duration'])
					except:
						continue

					allcounts[str(count_id)] = countdic
					# add results to count details file
					create_output(countdic,  countout,  COUNTCOLUMNS)


					# Create dictionaries to store the raw observations and the 'row' summary information.
					# This is a little duplication but the purest definition of 'clean data' seperates observations from analysis.
					# The observations file 'bike_movement_observations.txt' is the cleanest data.
 
					movedic = {}
					movedic['site_id']  		= countdic['site_id']
					movedic['count_id'] 		= countdic['count_id']
					movedic['bin_duration'] 		= countdic['bin_duration'] 

					bin_totals = {}
					bin_totals['site_id'] 		= countdic['site_id']
					bin_totals['count_id'] 		= countdic['count_id']
					bin_totals['bin_duration']	= countdic['bin_duration']

					# Check if all all detailed count information available.
					# Deal with the older data where there is no breakdown - or both male and female results are stored in "M" columns
					if countdic['bin_duration'] == 120:
						print 'Old super tuesday count, go to summary row'						
						# TODO: Handle this use case
					elif countdic['gender_split'] == 'N':
						print 'Male and female data combined'
						# TODO: Handle this use case
 
					else:
						start 	 = currentsheet['movement_bin_row_range']['seven_am_to_nine_am']['start']
						finish	 = currentsheet['movement_bin_row_range']['seven_am_to_nine_am']['finish']
						# Site has detailed '15min bin' bike movement information.						
						# TODO Fix hardcoded 7am-9am timerange in the preference file. It is no good for upfield study.

						# Collect details from each row and column in the count observations record
						for k in range(start,finish):         
							row = k + count_details
							col = currentsheet['movement_bin_times']['bin_start']['col']
							movedic['bin_start'] = sheet.cell(row,col).value

							# Collect details from each row of observations in the count                                        
							for i in currentsheet['movement_bin_columns']:
								col = currentsheet['movement_bin_columns'][i]['col']
								move = sheet.cell(row,col).value
								try:
									move = int(move)
								except:
									continue
								movedic[i] = move

							for summarise in summary_calculations:
								fieldsum = 0
								for i in range(0,len(summary_calculations[summarise])):
									interestingfield = summary_calculations[summarise][i]
									try:
										fieldsum = fieldsum + movedic[interestingfield]
									except:
										continue
								bin_totals[summarise] = fieldsum 

							# Convert the bin start time from excel into HH:MM 24hr format
							try:                   
								bin_start_time			= float(movedic['bin_start']) 
								preformatted_start_time	= xldate_as_tuple(bin_start_time,workbook.datemode)
								formatted_time 		= time(*preformatted_start_time[3:5])
								long_start_time		= datetime.combine(formatted_date,formatted_time)

								# Use observation start time as an identifier for the row                    
								movedic['bin_start'] 	= long_start_time
								bin_totals['bin_start'] 	= long_start_time
							except:
								continue
							
									
							# Output the row of observations, and the row summaries to text file
							create_output(movedic, moveout, MOVECOLUMNS)
							create_output(bin_totals,  bintotalout, BINTOTALS)
							


