#Working with json file and openpyxl
#The custom program scrapes data from a custom json file and saves in a excel file
#script written by Ayan Upadhaya, contact: ayanu881@gmail.com

import json
import openpyxl
from openpyxl.styles import Font

#importing data from json file
with open('persons.json','r') as f:
	data=json.load(f)

#creating a blank workbook
wb=openpyxl.Workbook()
sheet=wb.active

#columns and their headings and font style
columns=['A','B','C','D']
headings=['name','age','job','city']
fontObj=Font(name="Arial",bold=True)

#setting up column dimension
for cols in columns:
	sheet.column_dimensions[cols].width=15

#seting up titles for columns
index=1
for i in range(len(columns)):
	sheet[columns[i]+str(index)]=headings[i].capitalize()
	sheet[columns[i]+str(index)].font=fontObj

#exporting all data to excel file sheet
trac=2
for i in data['persons']:
	index=0
	sheet[columns[index]+str(trac)]=i['name'].capitalize()
	sheet[columns[index+1]+str(trac)]=i['age']
	sheet[columns[index+2]+str(trac)]=i['job'].capitalize()
	sheet[columns[index+3]+str(trac)]=i['city'].capitalize()
	trac+=1


wb.save('scrape.xlsx')

print("Success")

#how to read json data in terminal
#for i in data['persons']:
#	print(f"{i['name']}\t{i['age']}\t{i['job']}\t\t{i['city']}")