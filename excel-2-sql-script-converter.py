# coding=utf-8
import codecs
import xlrd
import sys
import re
if len(sys.argv)<2:
	#print "please input xls file name"
	exit()
try:
	outfile = sys.argv[2]
except:
	outfile = 'result.sql'
wb = xlrd.open_workbook(sys.argv[1])
f = codecs.open(outfile, 'w', 'utf-8')
#print "set autocommit = 0;"
f.write("set autocommit = 0;\n" )
config_sh = wb.sheet_by_name(u'config')
sheets = config_sh.col_values(0) #取出第一列
for sheet in sheets:
	if not sheet.startswith('tbl_'): #只写tbl_开头的
		continue
	sh = wb.sheet_by_name(sheet)
	colums = sh.row_values(0)
	try:
		colums = colums[0: colums.index('')]
	except:
		pass
	colums_num = len(colums)
	#print "-- Start processing table "+sheet
	#f.write("-- Start processing table "+sheet+"\n")
	#print "DELETE FROM "+sheet
	f.write("delete from `"+sheet+"`;\n")
	for rownum in range(1, sh.nrows):
		rowvalues = sh.row_values(rownum)
		rowvalues = rowvalues[:colums_num]
		tmp = [i for i in colums if i]; #删除空字段
		if not tmp: #删除空字段后整个数组为空, 此行为空
			continue
		sql = "insert into `"+sheet+"` ("
		for col in colums:
			sql = sql + "`"+col+"`,"
		sql = sql[:-1]
		sql = sql + ") values ("
		tmp = [i for i in rowvalues if i]
		if not tmp: #全为空
			continue;
		for colnum in range(colums_num):
			if type(rowvalues[colnum]) is str or type(rowvalues[colnum]) is unicode:
				if not rowvalues[colnum]:
					sql = sql +"null,"
				else:
					sql = sql + "'" + (rowvalues[colnum]) + "',"
			else:
				if not rowvalues[colnum]:
					sql = sql +"'0',"
				else:
					tmp_value = str(float(rowvalues[colnum]))
					tmp_value = re.sub(r'\.0$', '', tmp_value)
					sql = sql + "'"+ tmp_value +"',"
		sql = sql[:-1]
		sql = sql + ");"
		#print sql+""
		f.write(sql+"\n");
	#print "-- End of processing table "+sheet
	#f.write("-- End of processing table "+sheet+"\n")
#print "commit;"
f.write("commit;")
