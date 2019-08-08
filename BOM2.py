import xlrd
import xlsxwriter
from xlutils.copy import copy
workbookin = xlrd.open_workbook('BOM.xlsx')
worksheet = workbookin.sheet_by_index(0)
workbookout = xlsxwriter.Workbook()
sheet = workbookout.add_sheet('BOM')
workbookout.save('clearedBOM.xlsx')


readrow=1
writerow=0
item=0
QTY=1
partNumber=2
BOMStructure=3
description=4
def skipwhilelevel(readrow,BOMLevel):
	readrow=readrow+1
	while len(worksheet.cell(readrow, item).value) > len(BOMLevel):
			readrow=readrow+1
			print("removed child of purchased assembly")
			if readrow >= worksheet.nrows:
				break
	return readrow

while not worksheet.cell(readrow, item).value == xlrd.empty_cell.value:
	

	#barn till köpt assembly
	if readrow+1 == None:
		sheet.write(writerow, 7,worksheet.cell(readrow, QTY).value)
		sheet.write(writerow, 6,worksheet.cell(readrow, partNumber).value)
		sheet.write(writerow, 8,worksheet.cell(readrow, BOMStructure).value)
		sheet.write(writerow, 5,worksheet.cell(readrow, description).value)
		writerow=writerow+1
		workbookout.save('clearedBOM.xlsx')
		break
	if worksheet.cell(readrow, BOMStructure).value == "Purchased" and len(worksheet.cell(readrow, item).value) < len(worksheet.cell(readrow+1, item).value):
		sheet.write(writerow, 7,worksheet.cell(readrow, QTY).value)
		sheet.write(writerow, 6,worksheet.cell(readrow, partNumber).value)
		sheet.write(writerow, 8,worksheet.cell(readrow, BOMStructure).value)
		sheet.write(writerow, 5,worksheet.cell(readrow, description).value)
		writerow=writerow+1
		workbookout.save('clearedBOM.xlsx') 
		BOMLevel=worksheet.cell(readrow, item).value
		readrow=skipwhilelevel(readrow,BOMLevel)
		if readrow >= worksheet.nrows:
				break
				
	elif worksheet.cell(readrow, partNumber).value.startswith("NA"):
		print("Removed designed assembly")
		readrow=readrow+1

	else:
		sheet.write(writerow, 7,worksheet.cell(readrow, QTY).value)
		sheet.write(writerow, 6,worksheet.cell(readrow, partNumber).value)
		sheet.write(writerow, 8,worksheet.cell(readrow, BOMStructure).value)
		sheet.write(writerow, 5,worksheet.cell(readrow, description).value)
		writerow=writerow+1
		workbookout.save('clearedBOM.xlsx')
		readrow=readrow+1
