import pandas as pd
import numpy as np
import openpyxl as xl
import xlrd
import sys
import re
#import seaborn

seaborn.set(style='ticks')

def print_full(x):
	pd.set_option('display.max_rows', 1000)
	pd.set_option('display.max_columns', 100000)
	print(x)
	pd.reset_option('display.max_rows')
	pd.reset_option('display.max_columns')


##  argv[1] Analysis Type
##  argv[2] Data file (numbers)
##  argv[3] Tagging file with second sheet named "tags"
##  argv[4] Name of the desired outputfile 
##  argv[5] Nuture Analysis
	

try:
	sys.argv[2]
	data = pd.read_excel(sys.argv[2],header=0,inplace=True)
	dropData = pd.read_excel(sys.argv[2],header=0,sheet_name='Emails to Remove',inplace=True)
	for idx in dropData['Email ID']:
		data = data.drop(data[data['Email ID'] == idx].index) 
except:
	print('Please re-run and attach a xlsx  with Unique Opens, Email ID, Total Delivered, and Unique Clickthroughs')
	quit()


##Fill in the Blanks
data.fillna(0, inplace=True)
#splitSegments = data['Eloqua Segment'].str.extract('(BAB) Persona \w \S\s(\S.*)')
#Segments = splitSegments[0] + ' ' + splitSegments[1]
#data['Segments/Persona']= Segments
#analSegs = data[['Eloqua Segment','Segment/Persona','Email ID']].drop_duplicates()
#analSegs.fillna('General', inplace=True)
#print (analSegs)
#quit()

##Filter non-significant data
try:
	sys.argv[5].lower() == 'nurture'
	regex = re.compile('(BAB|RIA)')
	nurtureData = pd.read_excel(sys.argv[2],sheet_name='Nurture Emails',header=0,inplace=True)
	for idx in data['Email ID']:
		if idx in nurtureData['Email ID'].values:
			continue
		else:	
			data = data.drop(data[data['Email ID'] == idx].index)

except:
	print('Data is not Nurture')
	#data = data.drop(data[data['Total Delivered'] < 100].index)
	regex = re.compile('(BAB) Persona \w \S\s(\s|[^-]*)|(RIA)')

try:
	if 'Eloqua Segment' in data:
		
		splitSegments = data['Eloqua Segment'].str.extract(regex,expand=True)
		try:
			sys.argv[5].lower() == 'nurture'
#			print (splitSegments)
		
			
		except:
			splitSegments.fillna(' ', inplace=True)
			Segments = splitSegments.apply(lambda x: ' '.join(x.dropna().astype(str)),axis=1)
			Segments = Segments.str.replace('Established', 'Est.') 
			for idx, row in data.iterrows():
				if data.loc[idx,'Segment/Persona'] == 'BAB':
					if Segments[idx] != '     ' and Segments[idx] != 'RIA':
						data.loc[idx,'Segment/Persona'] = Segments[idx]
					else:
						data.loc[idx,'Segment/Persona'] = 'BAB'
		
				elif data.loc[idx,'Segment/Persona'] == 'RIA':
					continue
				else:
					data.loc[idx,'Segment/Persona'] = 'BAB'

		data['Segment/Persona'] = data['Segment/Persona'].str.strip()
#		print_full (data['Segment/Persona'])
		analSegs = data[['Segment/Persona','Email ID']].drop_duplicates()
#		print_full(analSegs)
	
except :
	print('No Eloqua Segments')
	pass


#Removing Duplicateus
data=data.groupby('Email ID').sum().reset_index()


#Filter non-significant data
try: 
	sys.argv[5].lower() != 'nuture'
except:
	print('Data has been filtered')
	data = data.drop(data[data['Total Delivered'] < 100].index)

##Removing Unnecessary Data to calculate click through rate and open rate
keep = ['Unique Opens', 'Email ID', 'Total Delivered', 'Unique Clickthroughs', 'Eloqua Segment', 'Segment/Persona']

for col in data.columns:
	if col in keep:
		continue
	else:
		data.drop(col, axis=1, inplace=True)


##Calculating Open rate and clickthrough rate
#Depreciated
#data['Open Rate']=data['Unique Opens']/data['Total Delivered']
#data['Click Through Rate']=data['Unique Clickthroughs']/data['Total Delivered']


##The tagging info from excel to read in
#tagData = pd.read_excel('SL_Tagging_201712_Final.xlsx',header=0,inplace=True)
try:
	sys.argv[3]
	tagData = pd.read_excel(sys.argv[3],header=0,inplace=True)
	try:
		if sys.argv[1].lower() == 'subjectline':
			tagData.drop(columns='Character Count', inplace=True)
	except:
		print ('Please specify analysis type and re-run')
		quit()
except:
	print('Please re-run and attach a Tagged xlsx with A Second Sheet titled "Tags"')
	quit()


try:
	data = pd.merge(analSegs,data,on='Email ID')
		
except:
	print('Eloque Segment Not Present')
	pass
#print_full(data)

#Creating Correlation Table
test = tagData.drop(columns=['Email ID','Email', 'Subject Line'])
test1 = test.groupby(list(test.columns.values),as_index=False)
test1 = test1.describe()
test1.drop(columns = ['unique'], level=1, inplace=True)
test1=test1.T.drop_duplicates().T


##Create a New DF to combine the numbers with the tags
mergeData = pd.merge(tagData,data,on='Email ID')
#mergeData.drop(columns='Character Count', inplace=True)

##Read In the Tags Used
book = xlrd.open_workbook(sys.argv[3])


#book = xlrd.open_workbook('TDA_Analysis_032015_112017_test.xlsx')
sheet = book.sheet_by_name('Tags')
tags = []
for c in range(sheet.ncols):
	for r in range(sheet.ncols):
		if sheet.cell_value(r,c):
			if r:
				tags.append((sheet.cell_value(0,c), sheet.cell_value(r,c)))


cols = pd.MultiIndex.from_tuples(tags)

try:
	sys.argv[1]
	if sys.argv[1].lower() == 'email':
		analType = 'Click Through Rate'
		rate='Unique Clickthroughs'
	else:
		analType = 'Open Rate'
		rate = 'Unique Opens'

except:
	print('Please specify an analysis and re-run')
	quit()


#analysis = pd.DataFrame(index=['Count', 'Average % '+ analType, 'Std'+ analType],columns=cols)
#<F11>analysis.fillna(0, inplace=True)
skipThese=['Email ID', 'Total Delivered', 'Unique Opens', 'Unique Clickthroughs','Open Rate', 'Click Through Rate', 'Email', 'Campaign', 'Subject Line','Character Count','Segment/Persona']

#Short For Separate Segement Analysis
SepSegAnal = {}

try:
	if 'Segment/Persona' in data:
		for seg in data['Segment/Persona'].unique():
			#if seg == 'Character Count':
			#	continue
			SepSegAnal[seg] = mergeData.loc[mergeData['Segment/Persona'] == seg, :]
			#print (SepSegAnal[seg])
		

except:
	print('No segment breakdown analysis')
	pass


SepSegAnal['Overall'] = mergeData



try:
	sys.argv[4]
	# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pd.ExcelWriter('/mnt/c/Users/bshaw/Documents/TD/' + sys.argv[4], engine='xlsxwriter')
except:
	print("All this time wasted because you forgot to specify an output file")
	quit()

rows = pd.MultiIndex.from_tuples([(df, metric) for df in SepSegAnal for metric in ['Count', 'Total Delivered', rate, 'Average % '+ analType, df+' Index']], names =('Segment', 'Metric'))

for df in SepSegAnal:
	#analysis = pd.DataFrame(index=['Count', 'Total Delivered', rate, 'Average % '+ analType, 'Std'+ analType],columns=cols)
	analysis = pd.DataFrame(index=['Count', 'Total Delivered', rate, 'Average % '+ analType, df+' Index'],columns=cols)
	analysis.drop('Character Count', axis=1, level=0, inplace=True)
	analysis.fillna(0, inplace=True)	

	for col in SepSegAnal[df].columns:
		if col in skipThese:
			continue
		else:
			attributes = pd.Series(SepSegAnal[df][col])
			countList = attributes.value_counts()
			ctrSum = SepSegAnal[df].groupby(col)[rate].sum()
			ctrDelivered = SepSegAnal[df].groupby(col)['Total Delivered'].sum()
			ctrMean = ctrSum.divide(ctrDelivered)
			#print (ctrMean)
			#quit()
			#ctrMean = SepSegAnal[df].groupby(col)[rate].mean()*100
			#ctrStd = SepSegAnal[df].groupby(col)[rate].std()
			for avg in ctrMean.index:
				#print (avg + ' : '+ str(ctrMean[avg]*100))	
				#analysis.loc['Sum',(col,avg)] = ctrSum[avg]
				analysis.loc['Average % '+ analType,(col,avg)] = ctrMean[avg]*100
				#analysis.loc['Std'+ analType,(col,avg)] = ctrStd[avg]
				analysis.loc['Total Delivered',(col,avg)] = ctrDelivered[avg]
				analysis.loc[rate,(col,avg)] = ctrSum[avg]
			for att in countList.index:
				analysis.loc['Count',(col,att)] = countList[att]
				#analysis.loc['Average %'+ analType,(col,att)] = ctrMean[avg]*100
	analysis.loc['Total Delivered', ('Totals','')] = SepSegAnal[df]['Total Delivered'].sum()
	analysis.loc[rate, ('Totals','')] = SepSegAnal[df][rate].sum()
	analysis.loc['Average % '+ analType, ('Totals','')] = analysis.loc[rate, ('Totals','')] / analysis.loc['Total Delivered', ('Totals','')]
	for col in analysis.columns:
		if col == ('Totals',''):
			analysis.loc[ df + ' Index', col] = 100
		#print(analysis.loc['Average % '+ analType, col] / analysis.loc['Index', ('Totals','')])	
		analysis.loc[ df + ' Index', col] = analysis.loc['Average % '+ analType, col] / analysis.loc['Average % '+ analType, ('Totals','')]
			
#			for idx in attributes:
#				analysis.loc['Index', (col, idx)] = analysis.loc['Average % '+ analType,(col,idx)] / analysis.loc['Total Delivered', 'Totals'] 

#Trying to plot
	#SepSegAnal[df].plot(kind='bar', subplots=True)
	corr = analysis[analysis.index != 'Count'].corr()
	# Add the first dataframe to the worksheet.
	analysis.to_excel(writer, sheet_name=df + ' Analysis')
	workbook = writer.book
	# Add a format. Red
	format1 = workbook.add_format({'bg_color': '#e60000'}) 
	# Add a format. Green fill
	format2 = workbook.add_format({'bg_color': '#198c19'})
	# Add a format. Light Red
	format3 = workbook.add_format({'bg_color': '#FFC7CE'})
	# Add a format. Light Green
	format4 = workbook.add_format({'bg_color': '#a3d1a3'})



	
	worksheet = writer.sheets[df + ' Analysis']
	if sys.argv[1].lower() == 'subjectline':
		worksheet.conditional_format('B8:AJ8', {'type': 'cell', 'criteria':'>=', 'value':120, 'format':format2})
		worksheet.conditional_format('B8:AJ8', {'type': 'cell', 'criteria':'<=', 'value':80, 'format':format1})
		worksheet.conditional_format('B8:AJ8', {'type': 'cell', 'criteria':'between', 'minimum':80, 'maximum':95, 'format':format3})
		worksheet.conditional_format('B8:AJ8', {'type': 'cell', 'criteria':'between', 'minimum':105, 'maximum':120, 'format':format4})
	else:
		worksheet.conditional_format('B8:BL8', {'type': 'cell', 'criteria':'between', 'minimum':105, 'maximum':120, 'format':format4})
		worksheet.conditional_format('B8:BL8', {'type': 'cell', 'criteria':'>=', 'value':120, 'format':format2})
		worksheet.conditional_format('B8:BL8', {'type': 'cell', 'criteria':'<=', 'value':80, 'format':format1})
		worksheet.conditional_format('B8:BL8', {'type': 'cell', 'criteria':'between', 'minimum':80, 'maximum':95, 'format':format3})
	#corr.to_excel(writer, sheet_name=df + ' Correlation')



#test1.to_excel(writer, sheet_name='Frequencies')
writer.save()

##If converted corr to series use this one
#corr.to_excel('/mnt/c/Users/bshaw/Documents/TD/' + sys.argv[3], sheet_name='Correlation')
#Old Print Statement
#analysis.to_excel('/mnt/c/Users/bshaw/Documents/TD/' + sys.argv[3])
#if sys.argv[1].lower() == 'email':
#	analType = 'CTR'
#	for col in mergeData.columns:
#		if col in skipThese:
#	        	continue
#		else:
#			attributes = pd.Series(mergeData[col])
#			countList = attributes.value_counts()
#			#ctrSum = mergeData.groupby(col)['Click Through Rate'].sum()
#			ctrMean = mergeData.groupby(col)['Click Through Rate'].mean()
#			ctrStd = mergeData.groupby(col)['Click Through Rate'].std()	
#			
#			for avg in ctrMean.index:
#				#analysis.loc['Sum',(col,avg)] = ctrSum[avg]
#				analysis.loc['Average',(col,avg)] = ctrMean[avg]
#				analysis.loc['Std',(col,avg)] = ctrStd[avg]
#			for att in countList.index:
#				analysis.loc['Count',(col,att)] = countList[att]
#
#elif sys.argv[1].lower() == 'subjectline':
#	analType = 'OR'
#	for col in mergeData.columns:
#		if col in skipThese:
#			continue
#		else:
#			attributes = pd.Series(mergeData[col])
#			countList = attributes.value_counts()
#			#ctrSum = mergeData.groupby(col)['Open Rate'].sum()
#			ctrMean = mergeData.groupby(col)['Open Rate'].mean()*100
#			ctrStd = mergeData.groupby(col)['Open Rate'].std()*100
#
#			for avg in ctrMean.index:
#				#analysis.loc['Sum',(col,avg)] = ctrSum[avg]
#				analysis.loc['Average'+ analType,(col,avg)] = ctrMean[avg]
#				analysis.loc['Std'+ analType,(col,avg)] = ctrStd[avg]
#			for att in countList.index:
#				analysis.loc['Count',(col,att)] = countList[att]
#	analysis.drop(columns='Character Count', inplace=True)
#else:
#	print('Please specify an analysis and re-run')
#	quit()

#analysis.drop(columns='Character Count', inplace=True)
#print (analysis.T)

#corr = analysis[analysis.index != 'Count'].corr()
#corr = analysis.T['Count'].groupby(analysis.index).corr(analysis.T['Average'])
#corr = analysis.loc[analysis.index[0]]
#print (corr.reset_index().corr())
#quit()

#try:
#	sys.argv[4]
#	# Create a Pandas Excel writer using XlsxWriter as the engine.
#	writer = pd.ExcelWriter('/mnt/c/Users/bshaw/Documents/TD/' + sys.argv[4], engine='xlsxwriter')
#
#	# Add the first dataframe to the worksheet.
#	analysis.to_excel(writer, sheet_name='Analysis')
#	corr.to_excel(writer, sheet_name='Correlation')
#	test1.to_excel(writer, sheet_name='Frequencies')
#	##If converted corr to series use this one
#	#corr.to_excel('/mnt/c/Users/bshaw/Documents/TD/' + sys.argv[3], sheet_name='Correlation')
#	
#	##Old Print Statement
#	#analysis.to_excel('/mnt/c/Users/bshaw/Documents/TD/' + sys.argv[3])
#	writer.save()
#except:
#	print("All this time wasted because you forgot to specify an output file")
#	quit()
