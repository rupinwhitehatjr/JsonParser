import pandas as pd

import xlwt
from xlwt import Workbook
  
# Workbook is created
wb = Workbook()


curriculum="APT"
df = pd.read_json(curriculum+'.json')
  
# add_sheet is used to create sheet.
sheet = wb.add_sheet(curriculum)  
df=df["data"]
r = 0    
c = 0  

for row in df:
	print (r, row["classNumber"],row["name"])
	#a='{}\t{}\t{}\t{}\t\n'.format(row["classNumber"],row["name"], row["description"], row["course"]["courseDescription"] )
	
	sheet.write(r, c, row["classNumber"]) 
	sheet.write(r, c+1, row["name"]) 
	sheet.write(r, c+2, row["description"])   
	sheet.write(r, c+3, row["course"]["courseDescription"]) 
	sheet.write(r, c+4, row["isCapstone"])
	CN=row["classNumber"]
	classList=CN.split(sep=" ", maxsplit=3)
	print(classList)
	s=1
	for d in classList:
		sheet.write(r, c+4+s, d)
		s=s+1
	
	r=r+1 
	#f.write(a)
#f.close()
wb.save(curriculum+'.xls')