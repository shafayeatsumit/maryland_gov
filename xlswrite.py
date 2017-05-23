import xlsxwriter

row = 0
col = 0

workbook = xlsxwriter.Workbook('sumit.xls')
worksheet = workbook.add_worksheet()

for i in range(10):
	worksheet.write(row, col,i)
	col +=1

workbook.close()	