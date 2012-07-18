#coding: utf-8

from simple_report.core.cursor import AbstractCursor, AbstractCalculateNextCursor

class Cursor(AbstractCursor):
    """
    Специализированный курсор для XLS таблиц
    """

    def __init__(self, column=None, row=None):
        super(Cursor, self).__init__()

        self.column = column or (0, 0)
        self.row = row or (0, 0)

    def _test_value(self, value):
        super(Cursor, self)._test_value(value)

        # Координаты в XLS таблицах имеют вид
        # (3, 0). 3 - Номер столбца. Нумерация с нуля
        #         0 - Номер строки. Нумерация с нуля.
        assert isinstance(value[0], int)

class CalculateNextCursor(AbstractCalculateNextCursor):
    """
    """

    def get_next_column(self, current_column, end_col, begin_col):

        return current_column + end_col - begin_col + 1

    def get_first_column(self):
        # Колонки нумеруются и номер первой нуль.
        return 0

    def get_first_row(self):

        return 0