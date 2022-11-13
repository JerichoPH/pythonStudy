import argparse
import os
import sys
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill


class ExcelReader:
	_excel: openpyxl.workbook.Workbook = None
	_sheet: openpyxl.worksheet.worksheet.Worksheet = None
	_content: list[Any] = []
	_rows: list = []
	_title: list[str] = []
	_filename: str = ''
	_read_title_row_number: int = 0
	_original_row_number: int = 1
	_finished_row_number: int = 0
	
	def __init__(self, filename: str, read_title_row_number: int = None, original_row_number: int = None, finished_row_number: int = None):
		"""
		初始化
		:param filename: 文件名
		:type filename: str
		:param read_title_row_number: 设置读取表头行标
		:type read_title_row_number: int
		:param original_row_number: 设置读取起始行标
		:type original_row_number: int
		:param finished_row_number: 设置读取终止行标
		:type finished_row_number: int
		"""
		self._filename = filename
		if read_title_row_number is not None:
			self._read_title_row_number = read_title_row_number
		if original_row_number is not None:
			self._original_row_number = original_row_number
		if finished_row_number is not None:
			self._finished_row_number = finished_row_number
	
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
		self._rows = list(self._sheet.rows)[self.get_original_row_number():self.get_finished_row_number() if self.get_finished_row_number() else None]
		self._content = [[str(cell.value) for cell in row_datum] for row_datum in tuple(self._sheet.rows)[self.get_original_row_number():self.get_finished_row_number() if self.get_finished_row_number() else None]]
		return self
	
	def read_entire_sheet_by_first(self) -> __init__:
		"""
		读取第一个sheet
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self.set_sheet_by_index().read_title().read_rows()
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
		self.set_sheet_by_name(worksheet_name).read_title().read_rows()
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
	
	def set_sheet_by_index(self, worksheet_index: int = 0) -> __init__:
		"""
		通过索引设置工作表
		:param worksheet_index: 工作表索引
		:type worksheet_index: int
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._sheet = self._excel[self._excel.sheetnames[worksheet_index]]
		return self
	
	def set_sheet_by_name(self, worksheet_name: str) -> __init__:
		"""
		通过名称设置工作表
		:param worksheet_name: 工作表名称
		:type worksheet_name: str
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._sheet = self._excel[worksheet_name]
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
	
	def get_rows(self) -> list:
		"""
		获取原始行
		:return: 原始行列表
		:rtype: list
		"""
		return self._rows
	
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
		return self._original_row_number if self._original_row_number else 1
	
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
	
	def save(self, filename: str = None) -> __init__:
		"""
		保存文件
		:param filename: 文件名（如不填写则保存回原文件）
		:type filename: str
		:return: 本类对象
		:rtype: excelHelper.ExcelReader
		"""
		self._excel.save(filename if filename else self._filename)
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
	_location: str = None
	_font_name: str = None
	_font_color: str = None
	_font_size: float = None
	_font_underline: str = None
	_font_bold: bool = None
	_font_italic: bool = None
	_font_vert_align: str = None
	_font_outline: bool = None
	_font_shadow: bool = None
	_border_style: str = None
	_border_color: str = '000000'
	_border_top_style: str = None
	_border_top_color: str = '000000'
	_border_bottom_style: str = None
	_border_bottom_color: str = '000000'
	_border_left_style: str = None
	_border_left_color: str = '000000'
	_border_right_style: str = None
	_border_right_color: str = '000000'
	_fill_fg_color: str = None
	
	BORDER_STYLE_THIN = 'thin'
	BORDER_STYLE_DOTTED = 'dotted'
	BORDER_STYLE_THICK = 'thick'
	BORDER_STYLE_MEDIUM_DASHED = 'mediumDashed'
	BORDER_STYLE_DASH_DOT_DOT = 'dashDotDot'
	BORDER_STYLE_MEDIUM = 'medium'
	BORDER_STYLE_DASH_DOT = 'dashDot'
	BORDER_STYLE_SLANT_DASH_DOT = 'slantDashDot'
	BORDER_STYLE_DASHED = 'dashed',
	BORDER_STYLE_MEDIUM_DASH_DOT = 'mediumDashDot'
	BORDER_STYLE_MEDIUM_DASH_DOT_DOT = 'mediumDashDotDot'
	BORDER_STYLE_HAIR = 'hair'
	BORDER_STYLE_DOUBLE = 'double'
	
	def __init__(
			self,
			content='',
			location: str = None,
			font_name: str = None,
			font_color: str = None,
			font_size: float = None,
			font_underline: str = None,
			font_bold: bool = None,
			font_italic: bool = None,
			font_vert_align: str = None,
			font_outline: bool = None,
			font_shadow: bool = None,
			border_style: str = None,
			border_color: str = None,
			border_top_style: str = None,
			border_top_color: str = None,
			border_bottom_style: str = None,
			border_bottom_color: str = None,
			border_left_style: str = None,
			border_left_color: str = None,
			border_right_style: str = None,
			border_right_color: str = None,
			fill_fg_color: str = 'FFFFFF',
	):
		"""
		初始化
		:param content: 单元格内容
		:type content: Any
		:param location: 单元格所在坐标
		:type location: str
		:param font_name: 字体
		:type font_name: str
		:param font_color: 颜色
		:type font_color: str
		:param font_size: 字号
		:type font_size: float
		:param font_underline: 下划线
		:type font_underline: str
		:param font_bold: 粗体
		:type font_bold: bool
		:param font_italic: 斜体
		:type font_italic: bool
		:param font_vert_align: 对齐
		:type font_vert_align: str
		:param font_outline: 外框线
		:type font_outline: bool
		:param font_shadow: 阴影
		:type font_shadow: bool
		:param border_style: 边框样式
		:type border_style: str
		:param border_top_style: 上边框样式
		:type border_top_style: str
		:param border_bottom_style: 下边框样式
		:type border_bottom_style: str
		:param border_left_style: 左边框样式
		:type border_left_style: str
		:param border_right_style: 右边框样式
		:type border_right_style: str
		:param fill_fg_color: 填充背景色
		:type fill_fg_color: str
		"""
		self._location = location
		self._content = content
		if font_name is not None:
			self._font_name = font_name
		if font_color is not None:
			self._font_color = font_color
		if font_size is not None:
			self._font_size = font_size
		if font_underline is not None:
			self._font_underline = font_underline
		if font_bold is not None:
			self._font_bold = font_bold
		if font_italic is not None:
			self._font_italic = font_italic
		if font_vert_align is not None:
			self._font_vert_align = font_vert_align
		if font_outline is not None:
			self._font_outline = font_outline
		if font_shadow is not None:
			self._font_shadow = font_shadow
		if border_style is not None:
			self._border_style = border_style
		if border_color is not None:
			self._border_color = border_color
		if border_top_style is not None:
			self._border_top_style = border_top_style
		if border_top_color is not None:
			self._border_top_color = border_top_color
		if border_bottom_style is not None:
			self._border_bottom_style = border_bottom_style
		if border_bottom_color is not None:
			self._border_bottom_color = border_bottom_color
		if border_left_style is not None:
			self._border_left_style = border_left_style
		if border_left_color is not None:
			self._border_left_color = border_left_color
		if border_right_style is not None:
			self._border_right_style = border_right_style
		if border_right_color is not None:
			self._border_right_color = border_right_color
		if fill_fg_color is not None:
			self._fill_fg_color = fill_fg_color
	
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
	
	def get_font_name(self) -> str:
		"""
		获取当前字体
		:return: 字体名称
		:rtype: str
		"""
		return self._font_name
	
	def set_font_name(self, font_name: str, when: bool = True) -> __init__:
		"""
		设置字体
		:param font_name: 字体名称
		:type font_name: str
		:param when: 是否执行设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		if when is True:
			self._font_name = font_name
		return self
	
	def get_font_color(self) -> str:
		"""
		获取字体颜色
		:return: 字体颜色
		:rtype: str
		"""
		return self._font_color
	
	def set_font_color(self, font_color: str, when: bool = True) -> __init__:
		"""
		设置字体颜色
		:param font_color: 字体颜色
		:type font_color: str
		:param when: 是否设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		if when is True:
			self._font_color = font_color
		return self
	
	def get_font_size(self) -> float:
		"""
		获取字号
		:return: 字号
		:rtype: float
		"""
		return self._font_size
	
	def set_font_size(self, font_size: float, when: bool = True) -> __init__:
		"""
		设置字号
		:param font_size: 字号
		:type font_size: float
		:param when: 是否立即设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		if when is True:
			self._font_size = font_size
		return self
	
	def get_font_underline(self) -> str:
		"""
		获取字体下划线
		:return: 下划线
		:rtype: bool
		"""
		return self._font_underline
	
	def set_font_underline(self, font_underline: str, when: bool = True) -> __init__:
		"""
		设置字体下划线
		:param font_underline: 下划线
		:type font_underline: bool
		:param when: 是否立即设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		if when is True:
			self._font_underline = font_underline
		return self
	
	def get_font_bold(self) -> bool:
		"""
		获取字体粗体值
		:return: 字体粗体值
		:rtype: int
		"""
		return self._font_bold
	
	def set_font_bold(self, font_bold: bool, when: bool = True) -> __init__:
		"""
		设置字体粗体值
		:param font_bold: 粗体值
		:type font_bold:
		:param when: 是否立即设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		if when is True:
			self._font_bold = font_bold
		return self
	
	def get_font_italic(self) -> bool:
		"""
		获取字体倾斜
		:return: 字体是否倾斜
		:rtype: bool
		"""
		return self._font_italic
	
	def set_font_italic(self, font_italic: bool, when: bool = True) -> __init__:
		"""
		设置字体是否倾斜
		:param font_italic: 字体是否倾斜
		:type font_italic: bool
		:param when: 是否立即设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		if when is True:
			self._font_italic = font_italic
		return self
	
	def get_font_vert_align(self) -> str:
		"""
		获取字体对其方式
		:return: 字体对其方式
		:rtype: str
		"""
		return self._font_vert_align
	
	def set_font_vert_align(self, font_vert_align: str, when: bool = True) -> __init__:
		"""
		设置字体对其方式
		:param font_vert_align: 字体对齐方式
		:type font_vert_align: str
		:param when: 是否立即设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		if when is True:
			self._font_vert_align = font_vert_align
		return self
	
	def get_font_outline(self) -> bool:
		"""
		获取外边框线
		:return: 外边框线
		:rtype: bool
		"""
		return self._font_outline
	
	def set_font_outline(self, font_outline: bool, when: bool = True) -> __init__:
		"""
		设置字体外框线
		:param font_outline: 字体外框线
		:type font_outline: bool
		:param when: 是否立即设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		if when is True:
			self._font_outline = font_outline
		return self
	
	def get_font_shadow(self) -> bool:
		"""
		获取字体阴影
		:return: 字体阴影
		:rtype: bool
		"""
		return self._font_shadow
	
	def set_font_shadow(self, font_shadow: bool, when: bool = True) -> __init__:
		"""
		设置字体阴影
		:param font_shadow: 字体阴影
		:type font_shadow: bool
		:param when: 是否立即设置：True
		:type when: bool
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
	
	def get_font(self) -> openpyxl.styles.Font:
		"""
		获取字体样式
		:return: 字体样式
		:rtype: openpyxl.styles.Font
		"""
		return openpyxl.styles.Font(
			name=self.get_font_name(),
			color=self.get_font_color(),
			size=self.get_font_size(),
			underline=self.get_font_underline(),
			bold=self.get_font_bold(),
			italic=self.get_font_italic(),
			vertAlign=self.get_font_vert_align(),
			outline=self.get_font_outline(),
			shadow=self.get_font_shadow(),
		)
	
	def get_border_style(self) -> str:
		"""
		获取边框样式
		:return: 边框样式
		:rtype: str
		"""
		return self._border_style
	
	def set_border_style(self, border_style: str) -> __init__:
		"""
		设置边框样式
		:param border_style: 边框样式
		:type border_style: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_style = border_style
		return self
	
	def get_border_top_style(self) -> str:
		"""
		获取上边框样式
		:return: 上边框样式
		:rtype: str
		"""
		return self._border_top_style
	
	def set_border_top_style(self, border_top_style: str) -> __init__:
		"""
		设置上边框样式
		:param border_top_style: 上边框样式
		:type border_top_style: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_top_style = border_top_style
		return self
	
	def get_border_bottom_style(self) -> str:
		"""
		获取下边框样式
		:return: 下边框样式
		:rtype: str
		"""
		return self._border_bottom_style
	
	def set_border_bottom_style(self, border_bottom_style: str) -> __init__:
		"""
		设置下边框样式
		:param border_bottom_style: 下边框样式
		:type border_bottom_style: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_bottom_style = border_bottom_style
		return self
	
	def get_border_left_style(self) -> str:
		"""
		获取左边框样式
		:return: 左边框样式
		:rtype: str
		"""
		return self._border_left_style
	
	def set_border_left_style(self, border_left_style: str) -> __init__:
		"""
		设置左边框样式
		:param border_left_style: 左边框样式
		:type border_left_style: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_left_style = border_left_style
		return self
	
	def get_border_right_style(self) -> str:
		"""
		获取右边框样式
		:return: 右边框样式
		:rtype: str
		"""
		return self._border_right_style
	
	def set_border_right_style(self, border_right_style: str) -> __init__:
		"""
		设置右边框样式
		:param border_right_style: 右边框样式
		:type border_right_style: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_right_style = border_right_style
		return self
	
	def get_border_color(self) -> str:
		"""
		获取边框颜色
		:return: 边框颜色
		:rtype: str
		"""
		return self._border_style
	
	def set_border_color(self, border_color: str) -> __init__:
		"""
		设置边框颜色
		:param border_color: 边框颜色
		:type border_color: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_color = border_color
		return self
	
	def get_border_top_color(self) -> str:
		"""
		获取上边框颜色
		:return: 上边框颜色
		:rtype: str
		"""
		return self._border_top_color
	
	def set_border_top_color(self, border_top_color: str) -> __init__:
		"""
		设置上边框颜色
		:param border_top_color: 上边框颜色
		:type border_top_color: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_top_color = border_top_color
		return self
	
	def get_border_bottom_color(self) -> str:
		"""
		获取下边框颜色
		:return: 下边框颜色
		:rtype: str
		"""
		return self._border_bottom_color
	
	def set_border_bottom_color(self, border_bottom_color: str) -> __init__:
		"""
		设置下边框颜色
		:param border_bottom_color: 下边框颜色
		:type border_bottom_color: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_bottom_color = border_bottom_color
		return self
	
	def get_border_left_color(self) -> str:
		"""
		获取左边框颜色
		:return: 左边框颜色
		:rtype: str
		"""
		return self._border_left_color
	
	def set_border_left_color(self, border_left_color: str) -> __init__:
		"""
		设置左边框颜色
		:param border_left_color: 左边框颜色
		:type border_left_color: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_left_color = border_left_color
		return self
	
	def get_border_right_color(self) -> str:
		"""
		获取右边框颜色
		:return: 右边框颜色
		:rtype: str
		"""
		return self._border_right_color
	
	def set_border_right_color(self, border_right_color: str) -> __init__:
		"""
		设置右边框颜色
		:param border_right_color: 右边框颜色
		:type border_right_color: str
		:return: 本类对象
		:rtype: excelHelper.ExcelWriterCell
		"""
		self._border_right_color = border_right_color
		return self
	
	def get_border(self) -> openpyxl.styles.Border:
		"""
		获取边框样式
		:return: 边框样式
		:rtype:
		"""
		return openpyxl.styles.Border(
			top=openpyxl.styles.Side(style=self._border_top_style if self._border_top_style is not None else self._border_style, color=self._border_top_color if self._border_top_color is not None else self._border_color),
			bottom=openpyxl.styles.Side(style=self._border_bottom_style if self._border_bottom_style is not None else self._border_style, color=self._border_bottom_color if self._border_bottom_color is not None else self._border_color),
			left=openpyxl.styles.Side(style=self._border_left_style if self._border_left_style is not None else self._border_style, color=self._border_left_color if self._border_left_color is not None else self._border_color),
			right=openpyxl.styles.Side(style=self._border_right_style if self._border_right_style is not None else self._border_style, color=self._border_right_color if self._border_right_color is not None else self._border_color),
		)
	
	def get_fill_fg_color(self) -> str:
		"""
		获取填充色
		:return: 填充色
		:rtype: str
		"""
		return self._fill_fg_color
	
	def set_fill_fg_color(self, fill_fg_color) -> __init__:
		"""
		设置填充色
		:param fill_fg_color: 填充色
		:type fill_fg_color: str
		:return: 本类对象
		:rtype: str
		"""
		self._fill_fg_color = fill_fg_color
		return self
	
	def get_fill(self) -> openpyxl.styles.PatternFill:
		"""
		获取填充样式
		:return: 填充样式
		:rtype: openpyxl.styles.PatternFill
		"""
		return openpyxl.styles.PatternFill(patternType='solid', fgColor=self._fill_fg_color)


class ExcelWriterRow:
	_cells: dict = {}
	
	def __init__(self, cells: dict = None):
		if cells:
			self._cells = cells
	
	def get_cells(self) -> dict:
		"""
		获取单元格组
		:return: 若干单元格组成的字典
		:rtype: Dict[str:ExcelWriterCell]
		"""
		return self._cells
	
	def set_cells(self, cells: dict) -> __init__:
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
		self._sheet[excel_writer_cell.get_location()].font = excel_writer_cell.get_font()
		self._sheet[excel_writer_cell.get_location()].border = excel_writer_cell.get_border()
		self._sheet[excel_writer_cell.get_location()].fill = excel_writer_cell.get_fill()
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
	1、-T --operation_type：【三个可选参数reader、writer和update】分别控制演示读取、写入和修改原表三个功能
	2、-I --input_filename：用于演示读取功能时，读取的表名称
	4、-R --relative_path：是否使用相路径，默认：True
	"""
	parser = argparse.ArgumentParser()
	parser.description = 'excel读取文件工区（只支持2007版）'
	parser.add_argument('-T', '--operation_type', help='执行类型：reader、writer、update', type=str, default='')
	# parser.add_argument('-I', '--input_filename', help='传入文件名，用于reader读取的excel的参数', type=str, default='')
	# parser.add_argument('-R', '--relative_path', help='是否使用相对路径', type=bool, default=True)
	args = parser.parse_args()
	operation_type = args.operation_type
	# input_filename = args.input_filename
	# relative_path = args.relative_path
	
	input_filename = 'factories.xlsx'
	relative_path = True
	
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
			xlwt.set_sheet_name('测试表'). \
				add_cell(
				ExcelWriterCell(
					content='测试1',
					location='A1',
					font_name='宋体',
					font_size=18,
					font_color='FF0000'
				)
			). \
				add_cell(
				ExcelWriterCell(
					content='测试2',
					location='B2',
					border_style=
					ExcelWriterCell.BORDER_STYLE_DOTTED,
					border_color='00FFFF',
					border_left_style=ExcelWriterCell.BORDER_STYLE_MEDIUM_DASH_DOT,
					fill_fg_color='FFFF00',
				)
			). \
				save()
		
		# 设置一行单元格写入
		with ExcelWriter(filename=os.path.join(sys.path[0], 'test-2.xlsx') if relative_path else 'test-2.xlsx') as xlwt:
			pass
	
	elif operation_type == 'update':
		# 修改表格
		PRO_PRICES = {'大蒜': 1, '芹菜': 2, '芒果': 3, }
		with ExcelReader(filename=os.path.join(sys.path[0], '工作簿1.xlsx') if relative_path else '工作簿1.xlsx') as xlrd:
			for row in xlrd.set_sheet_by_index().read_rows().get_rows():
				if row[0].value in PRO_PRICES:
					cell = ExcelWriterCell(
						content=PRO_PRICES[row[0].value],
						font_name='宋体',
						font_size=18,
						font_bold=True,
						font_underline='single',
						font_color='00FF00',
						border_style=ExcelWriterCell.BORDER_STYLE_MEDIUM,
						border_color='FFFF00',
					)
					row[1].value = cell.get_content()
					row[1].font = cell.get_font()
					row[1].border = cell.get_border()
			
			xlrd.save()
