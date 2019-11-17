import PyPDF2, glob, os, csv, sys
from pandas import ExcelWriter
import pandas as pd

# Function to extract the pageContent ---------------------
def extractContent(pageNumber):
	pageObj = pdfReader.getPage(pageNumber)
	pageContent = pageObj.extractText()
	return pageContent
# ---------------------------------------------------------

# Function to get word index ------------------------------
def indexFunction(word, content):
	contentLen, wordLen = len(content), len(word)
	for i in range(contentLen):
		new_word = content[i:(i+wordLen)]
		if new_word == word:
			return(i)
			break
# ---------------------------------------------------------

# Function to get the range of pages of each unit ---------
def pagesFunction():
	print('Page function-----------------------------------')
	aPageStart, aPageEnd = [], []
	lookUp = 'Unit no.:'
	last_page = number_of_pages - 1

	# Loop through all the pages
	for pageNumber in range(last_page):
		pageContent = extractContent(pageNumber)

		# Get the number of starting and ending pages
		if lookUp in pageContent:
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
def get_value_function(pageContent, wordStart, wordEnd, max_len = 45):
	# wordStart and wordEnd are unique values, not lists
	posStart = pageContent.index(wordStart) + len(wordStart)
	newContent = pageContent[posStart:]

	posEnd = indexFunction(wordEnd, newContent)
	unitFeature = newContent[:posEnd].strip()

	# Check the lenght in order to avoid errors
	if len(unitFeature) < max_len:
		return unitFeature
	else:
		print('Not valid feature. Content is too long')

# First page function -------------------------------------
def fpFunction():
	inner_list, outter_list = [], []
	print('Starting first page function--------------------')
	for page, pageEnd in zip(aPageStart, aPageEnd):
		print('Looking at', page, 'page')
		pageContent = extractContent(page)
		print('\n')

		# Get line
		wordStart = 'Unit no.'
		wordEnd = 'Fecha'
		line = get_value_function(pageContent, wordStart, wordEnd)

		# Get AHU type
		for ahu in ahus:
			if ahu in pageContent:
				dv = ahu
				break

		# Get reference
		wordStart = 'Planta no.'
		wordEnd = 'Unit no.'
		ref = get_value_function(pageContent, wordStart, wordEnd)

		# Airflow
		airflow = get_value_function(pageContent, ')', 'm')

		inner_list = [page, pageEnd, line, dv, ref, airflow]
		outter_list.append(inner_list)


	'''
	if (len(df_line) == len(df_ahu)) and (len(df_line) == len(df_ref)):
		print('Page function completed!------------------------')
		return df_line, df_ahu, df_ref
	else:
		print('Check the function. Lens do not match!')
		sys.exit()
	print('First page function done------------------------')
	print('\n')
	'''
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
		inner_list.append(page + 1) # In order to show real page number

		# Extract page content
		pageContent = extractContent(page)
		print('Checking at page number', page+1)

		for wordStart, wordEnd in zip(aWordStart, aWordEnd):
			print('Looking for ', wordStart, 'and', wordEnd)

			# Work in starting and ending pairs, page by page
			if (wordStart in pageContent) and (wordEnd in pageContent):
				print('Found on page', page+1)
				unitFeature = get_value_function(pageContent, wordStart, wordEnd)
				inner_list.append(unitFeature)
			else:
				# Reset inner list
				inner_list = []
				print('No luck this time')
				print('\n')
				# Exit loop and go for the next page
				break

		# Check the lenght and append to the outter list
		if len(inner_list) == len(aWordStart) + 1:
			print('New entry for the outter list!')
			print('\n')
			outter_list.append(inner_list)
			# Reset inner list for next feature
			inner_list = []

	try:
		return outter_list
	except:
		print('No outter_list found!')
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

	name = 'Units.xlsx'
	writer = pd.ExcelWriter(name)
	df_units.to_excel(writer, index = False)
	writer.save()

	print(units)


	aWordStart = ['caudal de aire', 'húmedas)', 'Potencia', 'Velocidad (nominal)', 'Amperios']
	aWordEnd = ['m', 'Pa', 'kW', 'RPM', 'Tensión']
	columns = ['Page', 'Airflow', 'Static Press.', 'Power Consump.', 'RPM', 'Amperes']
	outter = extractFeatures(aWordStart, aWordEnd, 0, last_page)
	df = pd.DataFrame(outter, columns = columns)

	print('\n')

	'''
	new_inner, new_outter = [], []
	power_consump = []
	number_of_fans = []
	amperes = []

	for fan in outter:
		page = fan[1]
		val = fan[3]
		if 'total' in val:
			# Map the AHU unit
			number_of_fans.append(get_value_function(val, '(', 'x'))
			power_consump.append(val[10:])
			amperes.append(get_value_function(fan[5], 'x', 'A')
		else:


			number_of_fans.append(1)
			power_consump.append(val[7:])
			amperes.append(fan[5][:-1])
		print(amperes)

	'''
	#df_line, df_ahu, df_ref = fpFunction()
	#df_airflow, df_static, df_number, df_power, df_rpm, df_amperes = ecFunction()

	#print(df_line)

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
	'''

	# Export to Excel

	name = 'Fans Results.xlsx'
	writer = pd.ExcelWriter(name)
	df.to_excel(writer, index = False)
	writer.save()


	pdfFileObj.close()