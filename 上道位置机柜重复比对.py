import argparse
import json
import os
import sys
from datetime import time, datetime

import excelHelper
import stdoutHelper
from stdoutHelper import StdoutHelper

parser = argparse.ArgumentParser()
parser.description = '比对电子车间和检修车间上道位置'
parser.add_argument('-P', '--paragraph_name', help='地区名称', type=str, default='')
args = parser.parse_args()
paragraph_name = args.paragraph_name


def read_excel(file_name: str) -> dict:
    install_shelves = {}

    with excelHelper.ExcelReader(file_name) as xlrd:
        excel_data = xlrd.set_sheet_by_index(0).set_original_row_number(2).auto_rows().to_dict

        for excel_datum in excel_data.values():
            install_shelves[excel_datum['架编码']] = excel_datum['名称']

        print(f'读取：', *stdoutHelper.content_info(file_name), *stdoutHelper.content_success('完成'))

    return install_shelves


if __name__ == '__main__':
    # 读取电子车间Excel
    ele_install_shelves = read_excel(os.path.join(sys.path[0], f'{paragraph_name}-电子车间-上道位置机柜.xlsx'))

    # 读取检修车间Excel
    fix_install_shelves = read_excel(os.path.join(sys.path[0], f'{paragraph_name}-检修车间-上道位置机柜.xlsx'))

    # 编号重复开始对比
    stdoutHelper.println_comment('编号重复开始对比')
    filename = os.path.join(sys.path[0], f'{paragraph_name}-重复-上道位置机柜.xlsx')
    repeat_unique_codes = set(ele_install_shelves.keys()).intersection(fix_install_shelves.keys())

    # 名称重复开始比对
    stdoutHelper.println_comment('名称重复开始比对')
    repeat_full_names = set(ele_install_shelves.values()).intersection(fix_install_shelves.values())
    fix_install_shelves_flip = {v: k for k, v in fix_install_shelves.items()}

    # 记录重复的编号
    repeats = []
    for _,datum in enumerate(repeat_unique_codes):
        repeats.append(datum)

    for _,datum in enumerate(repeat_full_names):
        if datum in fix_install_shelves_flip:
            repeats.append(fix_install_shelves_flip[datum])

    repeats = list(set(repeats))

    # 写入Excel
    with excelHelper.ExcelWriter(filename) as xlwt:
        xlwt.add_row(excelHelper.ExcelWriterRow(1, [
            excelHelper.ExcelWriterCell(content='机柜名称'),
            excelHelper.ExcelWriterCell(content='机柜代码'),
        ]))
        row_index = 1

        # 写入重复
        for unique_code in repeats:
            row_index += 1
            xlwt.add_row(excelHelper.ExcelWriterRow(row_index, [
                excelHelper.ExcelWriterCell(content=fix_install_shelves[unique_code] if unique_code in fix_install_shelves else ''),
                excelHelper.ExcelWriterCell(content=unique_code),
            ]))

        xlwt.save()

    stdoutHelper.println_success(f'比对完成，保存到{filename}')
