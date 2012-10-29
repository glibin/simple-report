#coding: utf-8
__author__ = 'prefer'

from simple_report.core.cursor import AbstractCursor, AbstractCalculateNextCursor
from simple_report.utils import ColumnHelper

class Cursor(AbstractCursor):
    """
    Специализированный курсор для XLSX таблиц.
    """

    def __init__(self, column=None, row=None, ):
        super(Cursor, self).__init__()
        self._column = column or ('A', 1)
        self._row = row or ('A', 1)

    def _test_value(self, value):

        super(Cursor, self)._test_value(value)

        # Координаты в XLSX таблицах имеют вид
        # (F, 3). F - имя стобла
        #         3 - номер строки. Нумерация строк с 1
        assert isinstance(value[0], basestring)


class CalculateNextCursorXLSX(AbstractCalculateNextCursor):
    """
    """

    def get_next_column(self, current_col, end_col, begin_col):

        return ColumnHelper.add(current_col, ColumnHelper.difference(end_col, begin_col) + 1)

    def get_first_column(self):
        # Колонки имеют строкое представление
        return 'A'

    def get_first_row(self):
        # Строки имеют числовое представление и нумер. с единицы.
        return 1

    def calculate_indent(self, column, w):
        """
        """
        return ColumnHelper.number_to_column(ColumnHelper.column_to_number(column)-
                                      w)
