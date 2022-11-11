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


if __name__ == '__main__':
	filename = 'factories.xlsx'
	
	with ExcelReader(filename=filename) as xlrd:
		# 通过Sheet名称进行读取
		excel_content = xlrd.read_entire_sheet_by_name('Sheet1').to_dict
		print('EX1：', excel_content)
		
		# 读取第一个Sheet
		excel_content = xlrd.read_entire_sheet_by_first().to_dict
		print('EX2：', excel_content)
		
		# 通过激活Sheet进行读取
		excel_content = xlrd.read_entire_sheet_by_active().to_dict
		print('EX3：', excel_content)
		
		# 读取表头
		excel_title = xlrd.set_read_title_row_number(0).set_title(xlrd.get_excel().active).read_title().get_title()
		print('EX4：', excel_title)
		
		# 读取多行数据
		excel_content = xlrd.set_sheet(xlrd.get_excel().active).set_original_row_number(2).set_finished_row_number(5).set_title(excel_title).read_rows().to_dict
		print('EX5：', excel_content)
