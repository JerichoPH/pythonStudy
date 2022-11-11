import openpyxl


class ExcelReader:
    _excel = None
    _sheet = None
    _content = []
    _titles = ()
    _filename = None
    _original_row_number = 2
    _finished_row_number = 0

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

    def read_title(self, row_number: int = 0) -> __init__:
        """
        读取一行数据
        :param row_number: 行号
        :type row_number: int
        :return: 本类对象
        :rtype: excelReader.ExcelReader.ExcelReader
        """
        self._titles = [str(cell.value) for cell in tuple(self._sheet.rows)[row_number]]
        return self

    def read_rows(self, original_row_number: int, finished_row_number: int = None) -> __init__:
        """
        读取多行数据
        :param original_row_number: 起始行号
        :type original_row_number: int
        :param finished_row_number: 终止行号
        :type finished_row_number: int
        :return: 本类对象
        :rtype: excelHelper.ExcelReader
        """
        self._content = [[str(cell.value) for cell in row] for row in tuple(self._sheet.rows)[self._original_row_number:self._finished_row_number]]
        return self

    def read_entire_sheet_by_first(self) -> __init__:
        """
        读取第一个sheet
        :return: 本类对象
        :rtype: excelHelper.ExcelReader
        """
        self.set_sheet(self._excel[self._excel.sheetnames[0]]).read_title().read_rows(original_row_number=2)
        return self

    def read_entire_sheet_by_active(self) -> __init__:
        """
        读取激活的sheet
        :return: 本类对象
        :rtype: excelHelper.ExcelReader
        """
        self.set_sheet(self._excel.active).read_title().read_rows(original_row_number=2)
        return self

    def read_entire_sheet_by_name(self, worksheet_name: str) -> __init__:
        """
        根据sheet名称读取
        :param worksheet_name: sheet名称
        :type worksheet_name: str
        :return: 本类对象
        :rtype: excelHelper.ExcelReader
        """
        self.set_sheet(self._excel[worksheet_name]).read_title().read_rows(original_row_number=2)
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

    def get_titles(self) -> list:
        """
        获取表头
        :return: 表头
        :rtype: tuple
        """
        return self._titles

    def set_titles(self, titles: list) -> __init__:
        """
        设置表头
        :param titles: 表头
        :type titles: tuple
        :return: 本类对象
        :rtype: excelHelper.ExcelReader
        """
        self._titles = titles
        return self

    def set_original_row_number(self, number: int) -> __init__:
        """
        设置起始行
        :param number: 起始行
        :type number: int
        :return: 本类对象
        :rtype: excelHelper.ExcelReader
        """
        pass

    @property
    def to_dict(self) -> list:
        """
        返回字典格式数据（必须设置表头）
        :return: Excel数据（带表头）
        :rtype: list
        """
        if self._titles:
            return [dict(zip(self._titles, content)) for content in self._content]

    @property
    def to_list(self) -> list:
        """
        返回数组格式数据
        :return:
        :rtype:
        """
        return self._content
