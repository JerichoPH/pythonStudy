import argparse
import json
import os
import sys
from datetime import time, datetime

import excelHelper
import stdoutHelper
from stdoutHelper import StdoutHelper

parser = argparse.ArgumentParser()
parser.description = '比对检修车间和电子车间重复器材'
parser.add_argument('-P', '--paragraph_name', help='地区名称', type=str, default='')
args = parser.parse_args()
paragraph_name = args.paragraph_name

if __name__ == '__main__':
    filename = os.path.join(sys.path[0], f'{paragraph_name}-电子车间-器材.xlsx')
    with excelHelper.ExcelReader(filename) as xlrd:
        ele_identity_codes_to_row_number = {}
        ele_data = xlrd.set_sheet_by_index(0).set_original_row_number(2).read_title().read_rows().to_dict
        ele_identity_codes = [datum['唯一编号'] for datum in ele_data.values()]
        for row_datum in ele_data.values():
            ele_identity_codes_to_row_number.setdefault(row_datum['唯一编号'], row_datum)
        print(f'读取：', *stdoutHelper.content_info(filename), *stdoutHelper.content_success('完成'))

    filename = os.path.join(sys.path[0], f'{paragraph_name}-检修车间-器材.xlsx')
    with excelHelper.ExcelReader(filename) as xlrd:
        fix_identity_codes_to_row_number = {}
        fix_data = xlrd.set_sheet_by_index(0).set_original_row_number(2).read_title().read_rows().to_dict
        fix_identity_codes = [datum['唯一编号'] for datum in fix_data.values()]
        for row_datum in fix_data.values():
            fix_identity_codes_to_row_number.setdefault(row_datum['唯一编号'], row_datum)
        print(f'读取：', *stdoutHelper.content_info(filename), *stdoutHelper.content_success('完成'))

    # 取两者交集
    stdoutHelper.print_comment('开始对比')
    filename = os.path.join(sys.path[0], f'{paragraph_name}-重复-器材.xlsx')
    intersection = set(ele_identity_codes).intersection(fix_identity_codes)

    with excelHelper.ExcelWriter(filename) as xlwt:
        xlwt.add_row(excelHelper.ExcelWriterRow(1, [
            excelHelper.ExcelWriterCell(content='唯一编号'),
            excelHelper.ExcelWriterCell(content='检修车间种类'),
            excelHelper.ExcelWriterCell(content='检修车间类型'),
            excelHelper.ExcelWriterCell(content='检修车间型号'),
            excelHelper.ExcelWriterCell(content='电子车间种类'),
            excelHelper.ExcelWriterCell(content='电子车间类型'),
            excelHelper.ExcelWriterCell(content='电子车间型号'),
        ]))

        for row_index, identity_code in enumerate(intersection):
            xlwt.add_row(excelHelper.ExcelWriterRow(row_index + 2, [
                excelHelper.ExcelWriterCell(content=identity_code),
                excelHelper.ExcelWriterCell(content=fix_identity_codes_to_row_number[identity_code]['种类'] if identity_code in fix_identity_codes_to_row_number else ''),
                excelHelper.ExcelWriterCell(content=fix_identity_codes_to_row_number[identity_code]['类型'] if identity_code in fix_identity_codes_to_row_number else ''),
                excelHelper.ExcelWriterCell(content=fix_identity_codes_to_row_number[identity_code]['型号'] if identity_code in fix_identity_codes_to_row_number else ''),
                excelHelper.ExcelWriterCell(content=ele_identity_codes_to_row_number[identity_code]['种类'] if identity_code in ele_identity_codes_to_row_number else ''),
                excelHelper.ExcelWriterCell(content=ele_identity_codes_to_row_number[identity_code]['类型'] if identity_code in ele_identity_codes_to_row_number else ''),
                excelHelper.ExcelWriterCell(content=ele_identity_codes_to_row_number[identity_code]['型号'] if identity_code in ele_identity_codes_to_row_number else ''),
            ]))
        xlwt.save()
    stdoutHelper.println_success(f'比对完成，保存到{filename}')
