from typing import Any


def print_wrong(content: Any, style: str = None) -> None:
	StdoutHelper.wrong(content, style).print()


def println_wrong(content: Any, style: str = None) -> None:
	StdoutHelper.wrong(content, style).println()


def print_warning(content: Any, style: str = None) -> None:
	StdoutHelper.warning(content, style).print()


def println_warning(content: Any, style: str = None) -> None:
	StdoutHelper.warning(content, style).println()


def print_info(content: Any, style: str = None) -> None:
	StdoutHelper.info(content, style).print()


def println_info(content: Any, style: str = None) -> None:
	StdoutHelper.info(content, style).println()


def print_comment(content: Any, style: str = None) -> None:
	StdoutHelper.comment(content, style).print()


def println_comment(content: Any, style: str = None) -> None:
	StdoutHelper.comment(content, style).println()


def print_success(content: Any, style: str = None) -> None:
	StdoutHelper.success(content, style).print()


def println_success(content: Any, style: str = None) -> None:
	StdoutHelper.success(content, style).println()


def print_normal(content: Any) -> None:
	StdoutHelper.normal(content).print()


def println_normal(content: Any) -> None:
	StdoutHelper.normal(content).println()


class StdoutHelper:
	_content = None
	_fg_color = ''  # 默认白色
	_bg_color = ''  # 默认黑色
	_style = ''  # 终端默认
	
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
	STYLE_HIGHLIGHT = '2'  # 高亮
	STYLE_ITALIC = '3'  # 倾斜
	STYLE_UNDERLINE = '4'  # 下横线
	STYLE_BLINK = '5'  # 闪烁
	STYLE_INVERSE = '7'  # 反白
	STYLE_INVISIBLE = '8'  # 不可见
	STYLE_DELETE_LINE = '9'  # 删除线
	
	def __init__(self, content: Any, fg_color: str = None, bg_color: str = None, style: str = None):
		self._content = content
		if fg_color is not None:
			self._fg_color = fg_color
		if bg_color is not None:
			self._bg_color = bg_color
		if style is not None:
			self._style = style
	
	@property
	def get_fg_color(self) -> str:
		"""
		获取前景色
		:return: 前景色
		:rtype: str
		"""
		return self._fg_color
	
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
	
	@property
	def get_bg_color(self) -> str:
		"""
		获取背景色
		:return: 背景色
		:rtype: str
		"""
		return self._bg_color
	
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
	
	@property
	def get_style(self) -> str:
		"""
		获取样式
		:return: 样式
		:rtype: str
		"""
		return self._style
	
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
	
	def println(self, sep=' ', end='\n') -> None:
		"""
		打印一行
		:return: 无
		:rtype: None
		"""
		print(*self.get_content, sep=sep, end=end)
	
	@property
	def get_original(self) -> str:
		"""
		获取起始标记
		:return: 起始标记
		:rtype: str
		"""
		return f'{self.ORIGINAL}{";".join(list(filter(None, [self.get_style, self.get_fg_color, self.get_bg_color])))}m'
	
	@property
	def get_finished(self) -> str:
		"""
		获取结束标记
		:return: 结束标记
		:rtype: str
		"""
		return self.FINISHED
	
	@property
	def get_content(self) -> list:
		"""
		返回需要打印的内容
		:return: 需要打印的内容
		:rtype: list
		"""
		return [self.get_original, self._content, self.get_finished]
	
	@staticmethod
	def wrong(content: Any, style: str = None):
		"""
		标记为错误
		:param content:
		:type content:
		:param style:
		:type style:
		:return:
		:rtype:
		"""
		ins = StdoutHelper(content).set_fg_color(StdoutHelper.COLOR_BLACK).set_bg_color(StdoutHelper.COLOR_RED)
		if style is not None:
			ins.set_style(style)
		return ins
	
	@staticmethod
	def warning(content: Any, style: str = None):
		"""
		标记为警告
		:param content:
		:type content:
		:param style:
		:type style:
		:return:
		:rtype:
		"""
		ins = StdoutHelper(content).set_fg_color(StdoutHelper.COLOR_BLACK).set_bg_color(StdoutHelper.COLOR_YELLOW)
		if style is not None:
			ins.set_style(style)
		return ins
	
	@staticmethod
	def success(content: Any, style: str = None):
		"""
		标记为成功
		:param content:
		:type content:
		:param style:
		:type style:
		:return:
		:rtype:
		"""
		ins = StdoutHelper(content).set_fg_color(StdoutHelper.COLOR_GREEN)
		if style is not None:
			ins.set_style(style)
		return ins
	
	@staticmethod
	def info(content: Any, style: str = None):
		"""
		标记为信息
		:param content:
		:type content:
		:param style:
		:type style:
		:return:
		:rtype:
		"""
		ins = StdoutHelper(content).set_fg_color(StdoutHelper.COLOR_CYAN)
		if style is not None:
			ins.set_style(style)
		return ins
	
	@staticmethod
	def comment(content: Any, style: str = None) -> __init__:
		"""
		标记为注释
		:param content:
		:type content:
		:param style:
		:type style:
		:return:
		:rtype:
		"""
		ins = StdoutHelper(content).set_fg_color(StdoutHelper.COLOR_PURPLE).set_style(StdoutHelper.STYLE_ITALIC)
		if style is not None:
			ins.set_style(style)
		return ins
	
	@staticmethod
	def normal(content: Any) -> __init__:
		"""
		普通消息
		:param content:
		:type content:
		:return:
		:rtype:
		"""
		return StdoutHelper(content)


if __name__ == '__main__':
	print('输出测试：\033[30m前景色 >> 30黑色\033[0m')
	StdoutHelper('前景色 >> 30黑色').set_fg_color(StdoutHelper.COLOR_BLACK).println()
	print('输出测试：\033[31m前景色 >> 31红色\033[0m')
	StdoutHelper('前景色 >> 31红色').set_fg_color(StdoutHelper.COLOR_RED).println()
	print('输出测试：\033[32m前景色 >> 32绿色\033[0m')
	StdoutHelper('前景色 >> 32绿色').set_fg_color(StdoutHelper.COLOR_GREEN).println()
	print('输出测试：\033[33m前景色 >> 33黄色\033[0m')
	StdoutHelper('前景色 >> 33黄色').set_fg_color(StdoutHelper.COLOR_YELLOW).println()
	print('输出测试：\033[34m前景色 >> 34蓝色\033[0m')
	StdoutHelper('前景色 >> 34蓝色').set_fg_color(StdoutHelper.COLOR_BLUE).println()
	print('输出测试：\033[35m前景色 >> 35紫色\033[0m')
	StdoutHelper('前景色 >> 35紫色').set_fg_color(StdoutHelper.COLOR_PURPLE).println()
	print('输出测试：\033[36m前景色 >> 36青色\033[0m')
	StdoutHelper('前景色 >> 36青色').set_fg_color(StdoutHelper.COLOR_CYAN).println()
	print('-' * 50)
	print('输出测试：\033[40m背景色 >> 40黑色\033[0m')
	StdoutHelper('背景色 >> 40黑色').set_bg_color(StdoutHelper.COLOR_BLACK).println()
	print('输出测试：\033[41m背景色 >> 41红色\033[0m')
	StdoutHelper('背景色 >> 41红色').set_bg_color(StdoutHelper.COLOR_RED).println()
	print('输出测试：\033[42m背景色 >> 42绿色\033[0m')
	StdoutHelper('背景色 >> 42绿色').set_bg_color(StdoutHelper.COLOR_GREEN).println()
	print('输出测试：\033[43m背景色 >> 43黄色\033[0m')
	StdoutHelper('背景色 >> 43黄色').set_bg_color(StdoutHelper.COLOR_YELLOW).println()
	print('输出测试：\033[44m背景色 >> 44蓝色\033[0m')
	StdoutHelper('背景色 >> 44蓝色').set_bg_color(StdoutHelper.COLOR_BLUE).println()
	print('输出测试：\033[45m背景色 >> 45紫色\033[0m')
	StdoutHelper('背景色 >> 45紫色').set_bg_color(StdoutHelper.COLOR_PURPLE).println()
	print('输出测试：\033[46m背景色 >> 46青色\033[0m')
	StdoutHelper('背景色 >> 46青色').set_bg_color(StdoutHelper.COLOR_CYAN).println()
	print('-' * 50)
	print('输出测试：\033[30;47m（前景色+背景色） >> 30黑色+47白色\033[0m')
	StdoutHelper('（前景色+背景色） >> 30黑色+47白色').set_bg_color(StdoutHelper.COLOR_WHITE).set_fg_color(StdoutHelper.COLOR_BLACK).println()
	print('输出测试：\033[30;41m（前景色+背景色） >> 30黑色+41红色\033[0m')
	StdoutHelper('（前景色+背景色） >> 30黑色+41红色').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_BLACK).println()
	print('-' * 50)
	print('输出测试：\033[0;30;41m样式 >> 0终端默认\033[0m')
	StdoutHelper('样式 >> 0终端默认').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_BLACK).set_style(StdoutHelper.STYLE_DEFAULT).println()
	print('输出测试：\033[1;30;41m样式 >> 1变暗\033[0m')
	StdoutHelper('样式 >> 1变暗').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_BLACK).set_style(StdoutHelper.STYLE_DARK).println()
	print('输出测试：\033[2;34;41m样式 >> 2高亮\033[0m')
	StdoutHelper('样式 >> 2高亮').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_BLUE).set_style(StdoutHelper.STYLE_HIGHLIGHT).println()
	print('输出测试：\033[3;34;41m样式 >> 3倾斜\033[0m')
	StdoutHelper('样式 >> 3倾斜').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_BLUE).set_style(StdoutHelper.STYLE_ITALIC).println()
	print('输出测试：\033[4;34;41m样式 >> 4下划线\033[0m')
	StdoutHelper('样式 >> 4下划线').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_BLUE).set_style(StdoutHelper.STYLE_UNDERLINE).println()
	print('输出测试：\033[5;37;41m样式 >> 5闪动\033[0m')
	StdoutHelper('样式 >> 5闪动').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_WHITE).set_style(StdoutHelper.STYLE_BLINK).println()
	print('输出测试：\033[7;37;41m样式 >> 7反白\033[0m')
	StdoutHelper('样式 >> 7反白').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_WHITE).set_style(StdoutHelper.STYLE_INVERSE).println()
	print('输出测试：\033[8;32;41m样式 >> 8隐藏\033[0m')
	StdoutHelper('样式 >> 8隐藏').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_BLACK).set_style(StdoutHelper.STYLE_INVISIBLE).println()
	print('输出测试：\033[9;37;41m样式 >> 9删除线\033[0m')
	StdoutHelper('样式 >> 9删除线').set_bg_color(StdoutHelper.COLOR_RED).set_fg_color(StdoutHelper.COLOR_WHITE).set_style(StdoutHelper.STYLE_DELETE_LINE).println()
	print('-' * 50)
	d = {'a': 'A', 'b': 'B'}
	l = [1, 2, 3]
	StdoutHelper(content=d).set_fg_color(StdoutHelper.COLOR_GREEN).set_bg_color(StdoutHelper.COLOR_WHITE).set_style(StdoutHelper.STYLE_BLINK).println()
	StdoutHelper(content=l, fg_color=StdoutHelper.COLOR_RED, bg_color=StdoutHelper.COLOR_WHITE, style=StdoutHelper.STYLE_ITALIC).println()
	StdoutHelper.wrong('错误').println()
	StdoutHelper.warning('警告').println()
	StdoutHelper.info('提醒').println()
	StdoutHelper.comment('注释').println()
	StdoutHelper.success('成功').println()
	StdoutHelper.normal('普通').println()
	println_warning('***' * 20 + '分割线' + '***' * 20)
	println_wrong('错误')
	println_warning('警告')
	println_info('提醒')
	println_comment('注释')
	println_success('成功')
	println_normal('普通')
	println_warning('***' * 20 + '分割线' + '***' * 20)
	print(*StdoutHelper.wrong('错误').get_content, *StdoutHelper.normal('+').get_content, *StdoutHelper.warning('警告').get_content)
