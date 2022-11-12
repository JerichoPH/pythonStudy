import argparse
import os
import sys
from typing import List, Any, Dict

import openpyxl


class ExcelReader:
	_excel: openpyxl.workbook.Workbook = None
	_sheet: openpyxl.worksheet.worksheet.Worksheet = None
	_content: list[Any] = []
	_title: list[str] = []
	_filename: str = ''
	_read_title_row_number: int = 0
	_original_row_number: int = 2
	_finished_row_number: int = 0
	
	def __init__(self, filename: str):
		self._filename = filename
	
	def __enter__(self) -> __init__:
		self._excel = openpyxl.load_workbook(self._filename)
		return self
	
	def __exit__(self, exc_type, exc_val, exc_tb) -> None:
		self._excel.close()
	
	def open_excel(self, filename: str) -> __init__:
		"""
		打开Excel文件
		:param filename: 文件路径
		:type filename: str
		:return: 本类对象
		:rtype: excelReader.ExcelReader.ExcelReader
		"""
		self._excel = openpyxl.load_workbook(filename)
		return self
	
	def read_title(self) -> __init__:
		"""
		读取一行数据
		:return: 本类对象
		:rtype: excelReader.ExcelReader.ExcelReader
		"""
		self._title = [str(cell.value) for cell in tuple(self._sheet.rows)[self.get_read_title_row_number()]]
		return self
	
	def read_rows(self) -> __init__:
		"""
		读取多行数据
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._content = [[str(cell.value) for cell in row_datum] for row_datum in tuple(self._sheet.rows)[self.get_original_row_number():self.get_finished_row_number() if self.get_finished_row_number() else None]]
		return self
	
	def read_entire_sheet_by_first(self) -> __init__:
		"""
		读取第一个sheet
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self.set_sheet(self._excel[self._excel.sheetnames[0]]).read_title().read_rows()
		return self
	
	def read_entire_sheet_by_active(self) -> __init__:
		"""
		读取激活的sheet
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self.set_sheet(self._excel.active).read_title().read_rows()
		return self
	
	def read_entire_sheet_by_name(self, worksheet_name: str) -> __init__:
		"""
		根据sheet名称读取
		:param worksheet_name: sheet名称
		:type worksheet_name: str
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self.set_sheet(self._excel[worksheet_name]).read_title().read_rows()
		return self
	
	def get_excel(self) -> openpyxl.workbook.Workbook:
		"""
		获取excel对象
		:return: Excel对象
		:rtype: openpyxl.workbook.Workbook
		"""
		return self._excel
	
	def get_sheet(self) -> openpyxl.worksheet.worksheet.Worksheet:
		"""
		获取sheet
		:return: Worksheet对象
		:rtype: openpyxl.worksheet.worksheet.Worksheet
		"""
		return self._sheet
	
	def set_sheet(self, worksheet: openpyxl.worksheet.worksheet.Worksheet) -> __init__:
		"""
		设置sheet
		:param worksheet: sheet对象
		:type worksheet: openpyxl.worksheet.worksheet.Worksheet
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._sheet = worksheet
		return self
	
	def get_title(self) -> list:
		"""
		获取表头
		:return: 表头
		:rtype: tuple
		"""
		return self._title
	
	def set_title(self, titles: list) -> __init__:
		"""
		设置表头
		:param titles: 表头
		:type titles: tuple
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._title = titles
		return self
	
	def get_read_title_row_number(self) -> int:
		"""
		获取读取表头行标
		:return: 行标
		:rtype: int
		"""
		return self._read_title_row_number if self._read_title_row_number else 0
	
	def set_read_title_row_number(self, row_number: int) -> __init__:
		"""
		设置读取表头行
		:param row_number: 行标
		:type row_number: int
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._read_title_row_number = row_number
		return self
	
	def get_original_row_number(self) -> int:
		"""
		获取起始读取行标
		:return: 行标
		:rtype: int
		"""
		return self._original_row_number if self._original_row_number else 2
	
	def set_original_row_number(self, now_number: int) -> __init__:
		"""
		设置起始读取行标
		:param now_number: 行标
		:type now_number: int
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._original_row_number = now_number
		return self
	
	def get_finished_row_number(self) -> int:
		return self._finished_row_number if self._finished_row_number else 0
	
	def set_finished_row_number(self, row_number: int) -> __init__:
		"""
		设置终止读取行标
		:param row_number: 行标
		:type row_number: int
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._finished_row_number = row_number + 1
		return self
	
	@property
	def to_dict(self) -> dict:
		"""
		返回字典格式数据（必须设置表头）
		:return: Excel数据（带表头）
		:rtype: dict
		"""
		if self._title:
			contents = [dict(zip(self._title, content)) for content in self._content]
			return {str(row_number + self.get_original_row_number()): row_datum for row_number, row_datum in enumerate(contents)}
		else:
			return {}
	
	@property
	def to_list(self) -> list:
		"""
		返回数组格式数据
		:return:
		:rtype:
		"""
		return self._content


class ExcelWriterCell:
	_content: str = ''
	_location: str = ''
	
	def __init__(self, location: str, content: str = ''):
		self._location = location
		self._content = content
	
	def get_location(self) -> str:
		"""
		返回单元格所在位置
		:return: 单元格坐标
		:rtype: str
		"""
		return self._location
	
	def set_location(self, location: str) -> __init__:
		"""
		设置单元格所在位置
		:param location: 单元格坐标
		:type location: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._location = location
		return self
	
	def get_content(self) -> str:
		"""
		获取单元格内容
		:return:
		:rtype:
		"""
		return self._content
	
	def set_content(self, content: str = '') -> __init__:
		"""
		设置单元格内容
		:param content: 单元格内容
		:type content:
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._content = content
		return self


class ExcelWriterRow:
	_cells: Dict[str:ExcelWriterCell] = {}
	
	def __init__(self, cells: Dict[str:ExcelWriterCell] = None):
		if cells:
			self._cells = cells
	
	def get_cells(self) -> Dict[str:ExcelWriterCell]:
		"""
		获取单元格组
		:return: 若干单元格组成的字典
		:rtype: Dict[str:ExcelWriterCell]
		"""
		return self._cells
	
	def set_cells(self, cells: Dict[str:ExcelWriterCell]) -> __init__:
		"""
		设置单元格组
		:param cells: 字典形式的单元格组合
		:type cells: Dict[str:ExcelWriterCell]
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterRow
		"""
		self._cells = cells
		return self


class ExcelWriter:
	_excel: openpyxl.Workbook = None
	_sheet: openpyxl.worksheet.worksheet.Worksheet = None
	_filename: str = ''
	
	def __init__(self, filename: str):
		"""
		初始化
		:param filename: 文件名
		:type filename: str
		"""
		self._excel = openpyxl.Workbook()
		self._filename = filename
		self._sheet = self._excel.active
	
	def __enter__(self) -> __init__:
		return self
	
	def __exit__(self, exc_type, exc_val, exc_tb) -> None:
		self._excel.close()
	
	def set_sheet_name(self, worksheet_name: str) -> __init__:
		"""
		设置工作表名称
		:param worksheet_name: 工作表名称
		:type worksheet_name:  str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriter
		"""
		self._sheet.title = worksheet_name
		return self
	
	def add_sheet(self, worksheet_name: str, worksheet_index: int = None) -> __init__:
		"""
		增加一个工作表
		:param worksheet_name: 工作簿名称
		:type worksheet_name: str
		:param worksheet_index: 工作簿所在位置
		:type worksheet_index: int
		:return: 本类对象
		:rtype: excelHelper.ExcelWriter
		"""
		self._excel.create_sheet(title=worksheet_name, index=worksheet_index)
		self._sheet = self._excel[worksheet_name]
		return self
	
	def del_sheet(self, del_worksheet_name: str, select_worksheet_name: str = '') -> __init__:
		"""
		删除一个工作表
		:param del_worksheet_name: 需要删除的工作表名称
		:type del_worksheet_name: str
		:param select_worksheet_name: 删除工作表之后，选中的工作表名称
		:type select_worksheet_name: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriter
		"""
		del self._excel[del_worksheet_name]
		if select_worksheet_name:
			self._sheet = self._excel[select_worksheet_name]
		else:
			self._sheet = None
		return self
	
	def add_cell(self, excel_writer_cell: ExcelWriterCell) -> __init__:
		"""
		添加单元格
		:param excel_writer_cell: 单元格对象
		:type excel_writer_cell: ExcelWriterCell
		:return: 本类对象
		:rtype: excelHelper.ExcelWriter
		"""
		self._sheet[excel_writer_cell.get_location()] = excel_writer_cell.get_content()
		return self
	
	def del_cell(self, location: str) -> __init__:
		"""
		删除一个单元格
		:param location: 单元格坐标
		:type location: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriter
		"""
		del self._sheet[location]
		return self
	
	def add_row(self, excel_writer_row: ExcelWriterRow) -> __init__:
		"""
		添加一行数据
		:param excel_writer_row: 添加一行数据
		:type excel_writer_row: ExcelWriterRow
		:return: 本类对象
		:rtype: excelHelper.ExcelWriter
		"""
		pass
	
	def save(self) -> __init__:
		"""
		保存文件
		:return:
		:rtype:
		"""
		self._excel.save(self._filename)
		return self


if __name__ == '__main__':
	"""
	参数说明
	1、-T --operation_type：【两个可选参数reader和writer】分别控制演示ExcelReader和ExcelWriter两个功能
	2、-I --input_filename：用于演示ExcelReader功能时，读取的表名称
	4、-R --relative_path：作用于【reader和writer】两个文件名是否使用相路径，默认：True
	"""
	parser = argparse.ArgumentParser()
	parser.description = 'excel读取文件工区（只支持2007版）'
	parser.add_argument('-T', '--operation_type', help='执行类型：reader、writer', type=str, default='')
	parser.add_argument('-I', '--input_filename', help='传入文件名，用于reader读取的excel的参数', type=str, default='')
	parser.add_argument('-R', '--relative_path', help='是否使用相对路径', type=bool, default=True)
	args = parser.parse_args()
	operation_type = args.operation_type
	input_filename = args.input_filename
	relative_path = args.relative_path
	
	if operation_type == 'reader':
		# ExcelReader演示
		with ExcelReader(filename=os.path.join(sys.path[0], input_filename) if relative_path else input_filename) as xlrd:
			# 通过工作表名称进行读取
			excel_content = xlrd.read_entire_sheet_by_name('Sheet1').to_dict
			print('example1：', excel_content)
			
			# 读取第一个Sheet
			excel_content = xlrd.read_entire_sheet_by_first().to_dict
			print('example：', excel_content)
			
			# 通过激活Sheet进行读取
			excel_content = xlrd.read_entire_sheet_by_active().to_dict
			print('example3：', excel_content)
			
			# 读取表头
			excel_title = xlrd.set_read_title_row_number(0).set_title(xlrd.get_excel().active).read_title().get_title()
			print('example：', excel_title)
			
			# 读取多行数据
			excel_content = xlrd.set_sheet(xlrd.get_excel().active).set_original_row_number(2).set_finished_row_number(5).set_title(excel_title).read_rows().to_dict
			print('example5：', excel_content)
			
			# 获取列表数据
			excel_content = xlrd.set_sheet(xlrd.get_excel().active).set_original_row_number(2).set_finished_row_number(5).set_title(excel_title).read_rows().to_list
			print('example6：', excel_content)
	
	elif operation_type == 'writer':
		# ExcelWriter演示
		# 设置单独单元格写入
		with ExcelWriter(filename=os.path.join(sys.path[0], 'test-1.xlsx') if relative_path else 'test-1.xlsx') as xlwt:
			xlwt.set_sheet_name('测试表').add_cell(ExcelWriterCell('A1', '测试1')).add_cell(ExcelWriterCell('B2', '测试2')).save()
		
		# 设置一行单元格写入
		with ExcelWriter(filename=os.path.join(sys.path[0], 'test-2.xlsx') if relative_path else 'test-2.xlsx') as xlwt:
			pass
