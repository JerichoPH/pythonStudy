import argparse
import json
import os
import sys
from datetime import time, datetime

import excelHelper

parser = argparse.ArgumentParser()
parser.description = 'excel读取文件工区（只支持2007版）'
parser.add_argument('-I', '--input_filename', help='传入文件名', type=str, default='')
parser.add_argument('-O', '--output_filename', help='传出文件名', type=str, default='')
parser.add_argument('-R', '--relative_path', help='是否使用相对路径', type=bool, default=True)
args = parser.parse_args()
input_filename = args.input_filename
output_filename = args.output_filename
relative_path = args.relative_path

if __name__ == '__main__':
    original_time = datetime.now()
    with excelHelper.ExcelReader(filename=os.path.join(sys.path[0], input_filename) if relative_path else input_filename) as xlrd:
        # excel_content = xlrd.read_entire_sheet_by_name(worksheet_name='Sheet2').to_dict
        excel_sheet = xlrd.get_excel()['Sheet2']
        xlrd.set_sheet(excel_sheet)
        xlrd.read_title()
        xlrd.read_rows(original_row_number=2, finished_row_number=5)
        excel_content = xlrd.to_dict

        with open(file=os.path.join(sys.path[0], output_filename) if relative_path else output_filename, mode='w') as f:
            f.write(json.dumps(obj=excel_content, ensure_ascii=False))

    print(excel_content)
