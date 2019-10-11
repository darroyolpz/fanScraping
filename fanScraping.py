# Release date: 10-10-2019

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

# Possible main--------------------------------------------
for page in range(pageStart, pageEnd):
	# Default values
	inner_list = []

	# Extract page content
	pageContent = extractContent(page)
	print('Checking at page number', page+1)

	for wordStart, wordEnd in zip(aWordStart, aWordEnd):
		print('Looking for ', wordStart, 'and', wordEnd)

		'''
		For now all the fan data is in the same page. Once it finds the pair,
		just extract the motherfucking features and go home
		'''	

		# Work in starting and ending pairs, page by page
		if (wordStart in pageContent) and (wordEnd in pageContent):
			print('Found on page', page+1)
			unitFeature = get_value_function(pageContent, wordStart, wordEnd)
			inner_list.append(unitFeature)
		else:
			print('No luck this time')
			print('\n')
			# Exit loop and go for the next page
			break

	# Check the lenght and append to the outter list
	if len(inner_list) == len(aWordStart):
		print('New entry for the outter list!')
		print('\n')
		outter_list.append(inner_list)
		inner_list = []

return outter_list
#----------------------------------------------------------

# First page function -------------------------------------
def fpFunction():
	print('Starting first page function--------------------')
	df_ahu, df_line, df_ref = [], [], []
	for page in aPageStart:
		print('Looking at', page, 'page')
		pageContent = extractContent(page)
		print('\n')

		# Get line
		wordStart = 'Unit no.'
		wordEnd = 'Fecha'
		value = get_value_function(page, wordStart, wordEnd)
		df_line.append(value)

		# Get AHU type
		for ahu in ahus:
			if ahu in pageContent:
				df_ahu.append(ahu)
				break

		# Get reference
		wordStart = 'Planta no.'
		wordEnd = 'Unit no.'

		value = get_value_function(page, wordStart, wordEnd)
		df_ref.append(value)


	print('Len df_line', len(df_line))
	print('Len df_ahu', len(df_ahu))
	print('Len df_ref', len(df_ref))
	print('\n')

	if (len(df_line) == len(df_ahu)) and (len(df_line) == len(df_ref)):
		print('Page function completed!------------------------')
		return df_line, df_ahu, df_ref
	else:
		print('Check the function. Lens do not match!')
		sys.exit()
	print('First page function done------------------------')
	print('\n')

# EC fan function -----------------------------------------
def ecFunction():
	print('Starting EC function----------------------------')
	df_ec = []
	# Keywords
	keyword = ', Plug-fan'
	aWordStart = ['caudal de aire', 'húmedas)', 'Potencia', 'Velocidad (nominal)', 'Amperios']
	aWordEnd = ['m', 'Pa', 'kW', 'RPM', 'Tensión']

	# Loop through all the lines, pages by page
	for startPage, endPage in zip(aPageStart, aPageEnd):
		# Page by page
		for page in range(startPage, endPage):
			print('Page number:', page + 1)

			# Extract content
			pageContent = extractContent(page)

			if keyword in pageContent:
				try:
					print(keyword, 'found in page', page + 1)
					inner = []
					for wordStart, wordEnd in zip(aWordStart, aWordEnd):
						print('Looking for starting word', wordStart)

						posStart = pageContent.index(wordStart) + len(wordStart)
						newContent = pageContent[posStart:]

						if wordEnd in newContent:
							posEnd = indexFunction(wordEnd, newContent)
							unitFeature = newContent[:posEnd]
							inner.append(unitFeature)
							print(wordEnd, 'found already!')
						else:
							pageContent = extractContent(page + 1)
							posEnd = indexFunction(wordEnd, newContent)
							unitFeature = newContent[:posEnd]
							inner.append(unitFeature)
							print(wordEnd, 'found already!')

					df_ec.append(inner)
					print(df_ec)
					print('\n')
				except:
					print('False positive')
					pass

	df_airflow, df_static, df_number, df_power, df_rpm, df_amperes = [], [], [], [], [], []

	# Cleaning out my closet ------------------------------
	for row in df_ec:
		# Airflow
		df_airflow.append(row[0])

		# Static pressure
		df_static.append(row[1])

		# Number of fans - Multiple fans
		if 'x' in row[2]:
			print(row)
			# Number of fans - Multiple fans
			wordStart = '('
			wordEnd = 'x'

			posStart = row[2].index(wordStart) + len(wordStart)
			newContent = row[2][posStart:]
			posEnd = indexFunction(wordEnd, newContent)
			unitFeature = newContent[:posEnd].strip()			
			df_number.append(unitFeature) # Number of fans - Multiple fans

			# For power - Multiple fans
			wordStart = 'x'

			posStart = row[2].index(wordStart) + len(wordStart)
			newContent = row[2][posStart:]
			unitFeature = newContent.strip()			
			df_power.append(unitFeature)

		elif 'x' not in row[2]:
			# Number of fans - Single fan
			df_number.append('1')

			# For power - Single fan
			# For power - Multiple fans
			wordStart = 'nominal'

			posStart = row[2].index(wordStart) + len(wordStart)
			newContent = row[2][posStart:]
			unitFeature = newContent.strip()			
			df_power.append(unitFeature)


		# RPM
		df_rpm.append(row[3].strip())

		# For amperes
		if 'x' in row[4]:
			print(row)
			# Number of fans - Multiple fans
			wordStart = ')'
			wordEnd = 'A'

			posStart = row[4].index(wordStart) + len(wordStart)
			newContent = row[4][posStart:]
			posEnd = indexFunction(wordEnd, newContent)
			unitFeature = newContent[:posEnd].strip()			
			df_amperes.append(unitFeature)
		elif 'x' not in row[4]:
			unitFeature = row[4][:-1].strip()			
			df_amperes.append(unitFeature)
	'''
	print(df_ec)
	print('\n')
	print('Airflow:', df_airflow)
	print('Static pressure:', df_static)
	print('Number of fans:', df_number)
	print('Power:', df_power)
	print('RPM:', df_rpm)
	print('Amperes:', df_amperes)
	'''
	if (len(df_airflow) == len(df_static)) and (len(df_airflow) == len(df_number)) and (len(df_airflow) == len(df_power)) and (len(df_airflow) == len(df_rpm)) and (len(df_airflow) == len(df_amperes)):
		return df_airflow, df_static, df_number, df_power, df_rpm, df_amperes
		print('Everything well returned..........................................................................................')
		print('\n')
		print('\n')
		print('\n')
	else:
		print('Check the function. Lens do not match!')
		dfs = [df_airflow, df_static, df_number, df_power, df_rpm, df_amperes]
		for df in dfs:
			print(len(df))
		sys.exit()

	

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
	print(aPageStart, aPageEnd)
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

	# Export to Excel
	name = 'Fans Results.xlsx'
	writer = pd.ExcelWriter(name)
	df.to_excel(writer, index = False)
	writer.save()

	'''

	pdfFileObj.close()

#-----------------------------------------