import PyPDF2, glob, os, csv, sys
from pandas import ExcelWriter
import pandas as pd

# Function to extract the pageContent ---------------------
def extractContent(pageNumber):
	pageObj = pdfReader.getPage(pageNumber)
	pageContent = pageObj.extractText()
	return pageContent
# ---------------------------------------------------------

# Function to get the range of pages of each unit ---------
'''
One of the first functions to run. It separates the units inside the
entire pdf and allows a cleaner scraping. 
It must have a keyword to separate one unit from the other: in our case, 'Unit no.:'.
It returns the first and
last page of each unit. Useful for future iterations.
'''
def pagesFunction(keyword = 'Unit no.:'):
	print('Page function-----------------------------------')
	aPageStart, aPageEnd = [], []
	last_page = number_of_pages - 1

	# Loop through all the pages
	for pageNumber in range(last_page):
		pageContent = extractContent(pageNumber)

		# Get the number of starting and ending pages
		if keyword in pageContent:
			if len(aPageStart) == 0:
				aPageStart.append(pageNumber)
			else:
				aPageEnd.append(pageNumber - 1)
				aPageStart.append(pageNumber)

	# Get the very last page
	aPageEnd.append(last_page)

	print('Page function completed!------------------------')
	print('\n')
	return aPageStart, aPageEnd

# Get SINGLE value function---------------------------------
def get_value_function(pageContent, wordStart, wordEnd, min_len = 1, max_len = 45):
	# wordStart and wordEnd are unique values, not lists
	posStart = pageContent.index(wordStart) + len(wordStart)
	newContent = pageContent[posStart:]

	#posEnd = indexFunction(wordEnd, newContent)
	posEnd = newContent.index(wordEnd)
	unitFeature = newContent[:posEnd].strip()

	# Check the lenght in order to avoid errors
	if (len(unitFeature) > min_len) & (len(unitFeature) < max_len):
		return unitFeature
	else:
		print('Not valid feature. Content is too long')
		return 'Error flag!'

# First page function -------------------------------------
def fpFunction():
	print('Starting first page function--------------------')
	inner_list, outter_list = [], []
	for page, pageEnd in zip(aPageStart, aPageEnd):
		print('Looking at', page, 'page')
		pageContent = extractContent(page)
		print('\n')

		# Get line
		wordStart = 'Unit no.'
		wordEnd = 'Fecha'
		line = get_value_function(pageContent, wordStart, wordEnd)

		# Reset ahu_value for each pageStart
		ahu_value = []

		# Get AHU type
		for ahu in ahus:
			if ahu in pageContent:
				ahu_value.append(ahu)

		# In case of DV10 or DV100, always get the longest one
		if len(ahu_value) == 1:
			ahu = ahu_value[0]
		elif len(ahu_value) > 1:
			print('Possible conflict!')
			final_ahu = ''
			for value in ahu_value:
				if len(value) > len(final_ahu):
					final_ahu = value
			ahu = final_ahu
			print('Final value:', ahu)
			print('\n')
		else:
			ahu = '---'

		# Get reference
		wordStart = 'Planta no.'
		wordEnd = 'Unit no.'
		ref = get_value_function(pageContent, wordStart, wordEnd)

		# Airflow
		wordStart = ')'
		wordEnd = 'm'
		airflow = get_value_function(pageContent, wordStart, wordEnd)

		inner_list = [page, pageEnd, line, ahu, ref, airflow]
		outter_list.append(inner_list)

	return outter_list

# Possible main--------------------------------------------
'''
In order to get the ID of each fan, just retun the page number.
Create a new function to match the page number and the unit.

One should call it and just pass the aWordStart and aWordEnd lists to
extract the features. It should be performed across the entire document
(pageStart = 0, pageEnd = last_page) but it's good to have such function
so that it can check unit by unit
'''
def extractFeatures(aWordStart, aWordEnd, pageStart, pageEnd):
	outter_list = []
	for page in range(pageStart, pageEnd):
		# Initiate the inner_list and get the page number
		inner_list = []
		
		# Extract page content
		pageContent = extractContent(page)
		print('Checking at page number', page+1)

		for wordStart, wordEnd in zip(aWordStart, aWordEnd):
			print('Looking for ', wordStart, 'and', wordEnd)

			# Work in starting and ending pairs, page by page
			if (wordStart in pageContent) and (wordEnd in pageContent):
				print('Found on page', page+1)
				unitFeature = get_value_function(pageContent, wordStart, wordEnd)

				if unitFeature == 'Error flag!':
					print('Error flag! Length not correct.')
					break
				else:
					inner_list.append(unitFeature)
			else:
				if len(inner_list) == 0:
					# Reset inner list
					inner_list = []
					print('No luck this time')
					print('\n')
					# Exit loop and go for the next page
					break
				elif len(inner_list) > 0:
					allowed_pages = 1
					pageContent = extractContent(page + 1)
					try:
						unitFeature = get_value_function(pageContent, wordStart, wordEnd)
						inner_list.append(unitFeature)
					except:
						print('No luck even in the next page')
						print('\n')
						# Exit loop and go for the next page
						break

		# Check the lenght and append to the outter list
		if len(inner_list) == len(aWordStart):
			print('New entry for the outter list!')
			print('\n')
			inner_list = [page + 1, *inner_list] # In order to show real page number
			outter_list.append(inner_list)
			# Reset inner list for next feature
			inner_list = []

	try:
		return outter_list
	except:
		print('No outter_list found!')
#----------------------------------------------------------

# Number of fans ------------------------------------------
def number_of_fans_function(field, wordStart, wordEnd):
	unitFeature = get_value_function(pageContent, wordStart, wordEnd)
	cleaned = field.str.slice()
	return number_of_fans

#----------------------------------------------------------

# Power consump. Cleaning ---------------------------------
def power_consump_cleaning(field, wordStart, wordEnd):
	unitFeature = get_value_function(pageContent, wordStart, wordEnd)
	cleaned = field.str.slice()

#----------------------------------------------------------

# Main ----------------------------------------------------
# Open file and read it
path = os.path.dirname(os.path.realpath(__file__))
num_files = len(glob.glob1(path,'*.pdf'))

extList = []
newList = []

for fileName in glob.glob('*.pdf'):
	# Initialize ----------------------------------------------
	aDVSize = []
	aDVLine = []
	aPageStart = []

	ahus = ['DV10', 'DV15', 'DV20', 'DV25', 'DV30', 'DV40', 'DV50',
	'DV60', 'DV80', 'DV100', 'DV120', 'DV150', 'DV190', 'DV240',
	'Geniox 10', 'Geniox 11', 'Geniox 12', 'Geniox 14', 'Geniox 16',
	'Geniox 18', 'Geniox 20', 'Geniox 22', 'Geniox 24', 'Geniox 27',
	'Geniox 29']

	# Read the pdf --------------------------------------------
	fileName = fileName[:-4]
	pdfFileObj = open(fileName + '.pdf', 'rb')
	pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

	# Get number of pages
	number_of_pages = pdfReader.getNumPages()
	print('Number of pages:', number_of_pages)
	print('\n')

	# Get the range of the pages ------------------------------
	aPageStart, aPageEnd = pagesFunction()
	last_page = aPageEnd[-1:][0]
	print(aPageStart, aPageEnd)
	print('\n')

	units = fpFunction()
	columns_units = ['Page start', 'Page end', 'Line', 'AHU', 'Ref', 'Airflow']
	df_units = pd.DataFrame(units, columns = columns_units)

	# To get an accurate reading in the Excel file
	df_units['Page start'] = df_units['Page start'] + 1
	df_units['Page end'] = df_units['Page end'] + 1

	name = 'Units.xlsx'
	writer = pd.ExcelWriter(name)
	df_units.to_excel(writer, index = False)
	writer.save()
	
	aWordStart = ['caudal de aire', 'húmedas)', 'Potencia', 'Velocidad (nominal)', 'Amperios']
	aWordEnd = ['m', 'Pa', 'kW', 'RPM', 'Tensión']
	columns = ['Page', 'Airflow', 'Static Press.', 'Power Consump.', 'RPM', 'Amperes']
	outter = extractFeatures(aWordStart, aWordEnd, 0, last_page)
	df = pd.DataFrame(outter, columns = columns)

	print('\n')

	name = 'Fans Results.xlsx'
	writer = pd.ExcelWriter(name)
	df.to_excel(writer, index = False)
	writer.save()

	pdfFileObj.close()

	'''
	# Dataframe
	df = pd.DataFrame()
	df['Line'] = df_line
	df['AHU'] = df_ahu
	df['Ref'] = df_ref
	df['Airflow'] = df_airflow
	df['Static pressure'] = df_static
	df['Number of fans'] = df_number
	df['Power'] = df_power
	df['rpm'] = df_rpm
	df['A'] = df_amperes
	columns = ['Line', 'Airflow', 'Static pressure', 'Number of fans', 'Power', 'rpm', 'A']
	for col in columns:
		df[col] = df[col].astype(float)
	# Export to Excel
	name = 'Fans Results.xlsx'
	writer = pd.ExcelWriter(name)
	df.to_excel(writer, index = False)
	writer.save()
	'''