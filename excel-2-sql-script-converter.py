# coding=utf-8
import re
import codecs
import xlrd
import sys
if len(sys.argv)<2:
	print "please input xls file name"
	exit()
try:
	outfile = sys.argv[2]
except:
	outfile = 'result.sql'
wb = xlrd.open_workbook(sys.argv[1])
f = codecs.open(outfile, 'w', 'utf-8')
print "set autocommit = 0;"
f.write("set autocommit = 0;\n" )
for sheet in wb.sheet_names():
#	if not sheet.startswith('tbl_'): #只写tbl_开头的
#		continue
	sh = wb.sheet_by_name(sheet)
	colums = sh.row_values(0)
	try:
		colums = colums[0: colums.index('')]
	except:
		pass
	colums_num = len(colums)
	print "-- Start processing table "+sheet
	f.write("-- Start processing table "+sheet+"\n")
	print "DELETE FROM "+sheet
	f.write("DELETE FROM "+sheet+";\n")
	for rownum in range(1, sh.nrows):
		rowvalues = sh.row_values(rownum)
		rowvalues = rowvalues[:colums_num]
		tmp = [i for i in colums if i]; #删除空字段
		if not tmp: #删除空字段后整个数组为空, 此行为空
			continue
		sql = "INSERT INTO "+sheet+" ("
		for col in colums:
			sql = sql + "`"+col+"`,"
		sql = sql[:-1]
		sql = sql + ") VALUES ("
		tmp = [i for i in rowvalues if i]
		if not tmp: #全为空
			continue;
		for colnum in range(colums_num):
			if type(rowvalues[colnum]) is str or type(rowvalues[colnum]) is unicode:
				if not rowvalues[colnum]:
					sql = sql +"NULL,"
				else:
					sql = sql + "\"" + (rowvalues[colnum]) + "\","
			else:
				if not rowvalues[colnum]:
					sql = sql +"0,"
				else:
					sql = sql + re.sub('\.0$', '', str(float(rowvalues[colnum])))+","
		sql = sql[:-1]
		sql = sql + ");"
		print sql+""
		f.write(sql+"\n");
	print "-- End of processing table "+sheet
	f.write("-- End of processing table "+sheet+"\n")
print "commit;"
f.write("commit;")
