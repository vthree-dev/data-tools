import sys
import datetime
import xlrd
import xlwt
import re

if len(sys.argv) != 2:
    print('Usage: {} [filename]'.format(sys.argv[0]))
    exit(0)
xls = xlrd.open_workbook(sys.argv[1])

if 'templates' not in xls.sheet_names():
    print('No template found')
    raise ValueError('Not error')

crule = re.compile(r'\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日')
namerule = re.compile(r'{\w+}')

temp_sheet = xls.sheet_by_name('templates')
#print(temp_sheet.nrows, temp_sheet.ncols)
template = [temp_sheet.row_values(i) for i in range(temp_sheet.nrows)]

key_list = [ key[1:-1]
    for row_data in template for key in row_data
        if isinstance(key, str) and namerule.match(key) and key != '{pass}' ]

value_dict = {}
collected_data = []

def match_template_row(temprow, row, id1, id2):
    global value_dict
    if len(row) < len(temprow):
        return False
    for i in range(len(temprow)):
        if namerule.match(temprow[i]):
            value_dict[temprow[i][1:-1]] = row[i]
        elif temprow[i].strip() != str(row[i]).strip():
            print('{}:{}'.format(id1, id2), temprow[i], row[i])
            return False
    return True

for sheet_name in xls.sheet_names():
    if sheet_name == 'templates':
        continue
    print('Processing', sheet_name)
    sheet = xls.sheet_by_name(sheet_name)
    i = 0
    while i < sheet.nrows:
        match_rows = 0
        for k in range(temp_sheet.nrows):
            if i + k >= sheet.nrows:
                match_rows = 0
                break
            if match_template_row(template[k], sheet.row_values(i + k), k, i+k):
                match_rows += 1
            else:
                value_dict = {}
                match_rows = 0
                break
        if match_rows:
            collected_data.append(value_dict)
            i += match_rows
        else:
            i += 1
        value_dict = {}

workbook = xlwt.Workbook(encoding = 'utf-8')
worksheet = workbook.add_sheet('Worksheet')

for j, key in enumerate(key_list):
    worksheet.write(0, j, key)
for i, item in enumerate(collected_data):
    for j, key in enumerate(key_list):
        if key == 'tabledate':
            res = crule.search(item[key])
            if res is None:
                raise ValueError('[Error] tabledate invalid: ' + str(item[key]))
            date = datetime.date(*map(int, [res.group(i) for i in range(1, 4)]))
            worksheet.write(i + 1, j, str(date))
        else:
            worksheet.write(i + 1, j, item[key])

workbook.save('Excel_Workbook.xls')

