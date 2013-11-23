import zipfile, xlrd, sys, os, shutil, itertools

path_to_dir = os.getcwd()
path_to_dir_tmp = path_to_dir+'/tmpextract/'

for path, subdirs, files in os.walk(path_to_dir):
        for file in files:
            if file.endswith('.zip') and '(' in file:
                with zipfile.ZipFile(os.path.join(path, file), 'r') as zfile:
                    for name in zfile.namelist():
                        if name.endswith('.csv'): print file, '<---------- CHECK THIS'
                        if (name.endswith('.xls') or name.endswith('.xlsx')) and '(' in name:
                            print 'extracting %s' %(name)
                            zfile.extract(name,path_to_dir_tmp)

f = open(path_to_dir_tmp+'/SENTEMAILS.txt','wb')
for path, subdirs, files in os.walk(path_to_dir_tmp):
        for file in files:
            if file.endswith('.xls') or file.endswith('.xlsx'):     
                workbook = xlrd.open_workbook(os.path.join(path, file))
                worksheets = workbook.sheet_names()
                for worksheet_name in worksheets:
                    print 'reading', file, worksheet_name
                    worksheet = workbook.sheet_by_name(worksheet_name)
                    num_cells = worksheet.ncols - 1
                    curr_cell = 0
                    num_rows = worksheet.nrows - 1
                    curr_row = 0
                    while curr_cell < num_cells:
                        cell_type = worksheet.cell_type(curr_row, curr_cell)
                        cell_value = worksheet.cell_value(curr_row, curr_cell)
                        #if cell_value.lower() == 'company name': 
                         #   print 'FOUND'
                        if cell_value == 'Email':
                            while curr_row < num_rows:
                                curr_row += 1
                                cell_value = worksheet.cell_value(curr_row, curr_cell)
                                f.write(cell_value.encode('utf8') + '\n')
                        curr_cell += 1
f.close()

input = open(path_to_dir_tmp+'/SENTEMAILS.txt', 'rb')
output = open(path_to_dir+'/'+path_to_dir.split('\\')[-1].split(' (')[0]+'_PreviouslySent.txt', 'wb')
output.write('email\n')
count = 0
for key,  group in itertools.groupby(sorted(input)):
    output.write(key)
    count += 1
input.close()
output.close()
shutil.rmtree(path_to_dir_tmp)
print count, 'emails previously sent. Press any key to exit.', 
raw_input()