import xlwt
from datetime import datetime
from xlrd import open_workbook
from xlutils.copy import copy
from pathlib import Path

def output(filename, sheet, row_num, person_name, is_present):
    file_path = f'attendance_files/{filename}{datetime.now().date()}.xls'  # Adjust the path as needed
    if Path(file_path).is_file():
        rb = open_workbook(file_path)
        book = copy(rb)
        sh = book.get_sheet(0)  # Get the first sheet
    else:
        book = xlwt.Workbook()
        sh = book.add_sheet(sheet)

    style_bold_red = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
                                 num_format_str='#,##0.00')
    style_date = xlwt.easyxf(num_format_str='D-MMM-YY')

    sh.write(0, 0, datetime.now().date(), style_date)
    sh.write(1, 0, 'Name', style_bold_red)
    sh.write(1, 1, 'Present', style_bold_red)

    sh.write(row_num + 1, 0, person_name)
    sh.write(row_num + 1, 1, is_present)

    fullname = f'{filename}{datetime.now().date()}.xls'
    book.save(file_path)
    return fullname
