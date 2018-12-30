#conding:utf-8
import xlrd
import xlwt
import sys
import argparse

reload(sys)
sys.setdefaultencoding('utf-8')

def main(old,new,sheet,column):
	try:
		data_old = xlrd.open_workbook(old)
		table_old = data_old.sheets()[int(sheet)]
		nrows_old = table_old.nrows
		cn_old = table_old.col_values(int(column), start_rowx=1, end_rowx=int(nrows_old))
		
		data_new = xlrd.open_workbook(new)
		table_new = data_new.sheets()[int(sheet)]
		nrows_new = table_new.nrows
		cn_new = table_new.col_values(int(column), start_rowx=1, end_rowx=int(nrows_new))

		result1 = set(cn_old).difference(set(cn_new))
		result2 = set(cn_new).difference(set(cn_old))

		if result1:
			t1 = ['delete:']
			for i in result1:
				t1.append(i.strip())
		else:
			t1 = ['delete:']

		if result2:
			t2 = ['increase']
			for i in result2:
				t2.append(i.strip())
		else:
			t2 = ['increase']

		workbook = xlwt.Workbook(encoding = 'utf-8')
		worksheet = workbook.add_sheet('compare_result')
		n = 0
		for i in t1:
			worksheet.write(n,0, label = i)
			n += 1

		n = 0
		for i in t2:
			worksheet.write(n,1, label = i)
			n += 1
		workbook.save('change.xls')
	except Exception,e:
		print e

if __name__ == '__main__':
	parse = argparse.ArgumentParser()
	parse.add_argument('-o','--old', help="Enter the filename1 that you want to compare")
	parse.add_argument('-n','--new', help="Enter the filename2 that you want to compare")
	parse.add_argument('-s','--sheet', help="Enter the sheet number that you want to compare. exp: sheet1 -s 0")
	parse.add_argument('-c','--column', help="Enter the column number that you want to compare")
	args = parse.parse_args()

	old = args.old
	new = args.new
	sheet = args.sheet
	column = args.column

	main(old,new,sheet,column)