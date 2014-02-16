import xlrd
import xlwt
import os

def data_write():
	os.chdir("/home/roy/downloads/bigface")
	wbk = xlwt.Workbook()
	new_sheet = wbk.add_sheet("sheet1")
	m = 0
	for i in os.listdir("/home/roy/downloads/bigface/"):
		i = i.strip("\n")
		table = xlrd.open_workbook(i)
		print i
		sheet = table.sheet_by_index(0)
		for r in range(sheet.nrows):
			for c in range(sheet.ncols):
				new_sheet.write(m,c,sheet.cell(r,c).value)
			m += 1
	wbk.save("111.xls")

if __name__=="__main__":
	data_write()
