from typing import Any


class StdoutHelper:
	_content = None
	_fg_color = '37'  # 默认白色
	_bg_color = '40'  # 默认黑色
	_style = '0'  # 终端默认
	
	ORIGINAL = '\033['
	FINISHED = '\033[0m'
	
	COLOR_BLACK = '0'
	COLOR_RED = '1'
	COLOR_GREEN = '2'
	COLOR_YELLOW = '3'
	COLOR_BLUE = '4'
	COLOR_PURPLE = '5'
	COLOR_CYAN = '6'
	COLOR_WHITE = '7'
	
	STYLE_DEFAULT = '0'  # 终端默认
	STYLE_DARK = '1'  # 变暗
	STYLE_LIGHT = '2'  # 高亮
	STYLE_ITALIC = '3'  # 倾斜
	STYLE_UNDERLINE = '4'  # 下横线
	STYLE_BLINK = '5'  # 闪烁
	STYLE_INVERSE = '7'  # 反白
	STYLE_INVISIBLE = '8'  # 不可见
	STYLE_DELETE_LINE = '9'  # 删除线
	
	def __init__(self, content: Any):
		self._content = content
	
	def set_fg_color(self, fg_color: str) -> __init__:
		"""
		设置前景色
		:param fg_color: 前景色颜色
		:type fg_color: str
		:return: 本类对象
		:rtype: StdoutHelper
		"""
		self._fg_color = f'3{fg_color}'
		return self
	
	def set_bg_color(self, bg_color: str) -> __init__:
		"""
		设置背景色
		:param bg_color: 背景色
		:type bg_color: str
		:return: 本类对象
		:rtype: StdoutHelper
		"""
		self._bg_color = f'4{bg_color}'
		return self
	
	def set_style(self, style: str) -> __init__:
		"""
		设置样式
		:param style: 样式
		:type style: str
		:return: 本类对象
		:rtype: StdoutHelper
		"""
		self._style = style
		return self
	
	def print(self, sep=' ', end='') -> None:
		"""
		打印内容
		:return: 无
		:rtype: None
		"""
		print(*self.get_content, sep=sep, end=end)
	
	def print_line(self, sep=' ', end='\n') -> None:
		"""
		打印一行
		:return: 无
		:rtype: None
		"""
		print(*self.get_content, sep=sep, end=end)
	
	@property
	def get_original(self) -> str:
		"""
		获取内容
		:return:
		:rtype:
		"""
		return f'{self.ORIGINAL}{self._style};{self._fg_color};{self._bg_color}m'
	
	@property
	def get_finished(self):
		return f'{self.FINISHED}'
	
	@property
	def get_content(self) -> list:
		"""
		返回需要打印的内容
		:return: 需要打印的内容
		:rtype: list
		"""
		return [self.get_original, self._content, self.get_finished]


if __name__ == '__main__':
	from stdoutHelper import StdoutHelper
	
	print('输出测试：\033[30m前景色 >> 30黑色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_fg_color(StdoutHelper.COLOR_BLACK).print_line()
	print('输出测试：\033[31m前景色 >> 31红色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_fg_color(StdoutHelper.COLOR_RED).print_line()
	print('输出测试：\033[32m前景色 >> 32绿色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_fg_color(StdoutHelper.COLOR_GREEN).print_line()
	print('输出测试：\033[33m前景色 >> 33黄色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_fg_color(StdoutHelper.COLOR_YELLOW).print_line()
	print('输出测试：\033[34m前景色 >> 34蓝色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_fg_color(StdoutHelper.COLOR_BLUE).print_line()
	print('输出测试：\033[35m前景色 >> 35紫色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_fg_color(StdoutHelper.COLOR_PURPLE).print_line()
	print('输出测试：\033[36m前景色 >> 36青色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_fg_color(StdoutHelper.COLOR_CYAN).print_line()
	print('-' * 50)
	print('输出测试：\033[40m背景色 >> 40黑色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_bg_color(StdoutHelper.COLOR_BLACK).print_line()
	print('输出测试：\033[41m背景色 >> 41红色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_bg_color(StdoutHelper.COLOR_RED).print_line()
	print('输出测试：\033[42m背景色 >> 42绿色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_bg_color(StdoutHelper.COLOR_GREEN).print_line()
	print('输出测试：\033[43m背景色 >> 43黄色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_bg_color(StdoutHelper.COLOR_YELLOW).print_line()
	print('输出测试：\033[44m背景色 >> 44蓝色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_bg_color(StdoutHelper.COLOR_BLUE).print_line()
	print('输出测试：\033[45m背景色 >> 45紫色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_bg_color(StdoutHelper.COLOR_PURPLE).print_line()
	print('输出测试：\033[46m背景色 >> 46青色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_bg_color(StdoutHelper.COLOR_CYAN).print_line()
	print('-' * 50)
	print('输出测试：\033[30;47m（前景色+背景色） >> 30黑色+47白色\033[0m')
	print('输出测试：\033[30;41m（前景色+背景色） >> 30黑色+41红色\033[0m')
	print('-' * 50)
	print('输出测试：\033[0;30;41m样式 >> 0终端默认\033[0m')
	print('输出测试：\033[1;30;41m样式 >> 1变暗\033[0m')
	print('输出测试：\033[2;34;41m样式 >> 2高亮\033[0m')
	print('输出测试：\033[3;34;41m样式 >> 3倾斜\033[0m')
	print('输出测试：\033[4;34;41m样式 >> 4下划线\033[0m')
	print('输出测试：\033[5;37;41m样式 >> 5闪动\033[0m')
	print('输出测试：\033[6;37;41m样式 >> 6未知\033[0m')
	print('输出测试：\033[7;37;41m样式 >> 7反白\033[0m')
	print('输出测试：\033[8;32;41m样式 >> 8无字\033[0m')
	print('输出测试：\033[9;37;41m样式 >> 9删除线\033[0m')
	print('-' * 50)
	d = {'a': 'A', 'b': 'B'}
	StdoutHelper(content=d).set_fg_color(StdoutHelper.COLOR_GREEN).set_bg_color(StdoutHelper.COLOR_WHITE).set_style(StdoutHelper.STYLE_BLINK).print()
