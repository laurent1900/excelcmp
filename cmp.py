#conding:utf-8
import xlrd
import xlwt
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

def main():
	try:
		data_old = xlrd.open_workbook('company.xlsx')
		table_old = data_old.sheets()[0]
		nrows_old = table_old.nrows
		cn_old = table_old.col_values(0, start_rowx=1, end_rowx=int(nrows_old))
		
		data_new = xlrd.open_workbook('company2.xlsx')
		table_new = data_new.sheets()[0]
		nrows_new = table_new.nrows
		cn_new = table_new.col_values(0, start_rowx=1, end_rowx=int(nrows_new))

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
		worksheet = workbook.add_sheet('sheet1')
		n = 0
		for i in t1:
			worksheet.write(n,0, label = i)
			n += 1
			if n == len(t1):
				break

		n = 0
		for i in t2:
			worksheet.write(n,1, label = i)
			n += 1
			if n == len(t1):
				break
		workbook.save('change.xls')
	except Exception,e:
		print e

if __name__ == '__main__':
	main()