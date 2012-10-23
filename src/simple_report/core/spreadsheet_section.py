#coding: utf-8

import abc

from simple_report.utils import ColumnHelper
from simple_report.interface import ISpreadsheetSection

class SpreadsheetSection(object):

    __metaclass__ = abc.ABCMeta

    def __init__(self, sheet, name, begin=None, end=None):
        """
        Абстракная секция для таблиц
        """

        assert begin or end

        self.name = name
        self.begin = begin
        self.end = end


class AbstractMerge(object):
    """
    Конструкция Merge
    """

    __metaclass__ = abc.ABCMeta

    def __init__(self, section, params, oriented=ISpreadsheetSection.LEFT_DOWN,
                 from_new_row=True):

        self.section = section
        self.params = params
        self.oriented = oriented

        self.from_new_row = from_new_row

        # Строка с которой начинаем обьединять ячейки
        self.begin_row_merge = self._get_border_row()

    def __enter__(self):
        self.section.flush(self.params, self.oriented)

        # Индекс колонки, которую мержим
        column, _ = self.section.sheet_data.cursor.column
        self._merge_col = self._calculate_merge_column(column)

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        """

        self.end_row_merge = self._get_border_row(top_border=False)
        self._merge()

    @abc.abstractmethod
    def _calculate_merge_column(self, column):
        """
        Вычисление столбца, строки которого будем мержить. По сути вернуть предыдущий столбец
        """

    @abc.abstractmethod
    def _merge(self):
        """
        """

    def _get_border_row(self, top_border=True):
        # Функция вычисляет и возвращает номер строки с которой необходимо
        # начать и закончить мержить.
        # Результат работы зависит от курсора.
        # Параметр begin указывает на то, какая граница вычисляется.
        # top_border = True - Верхняя граница

        # Колонка и строка курсора row
        _, r_row = self.section.sheet_data.cursor.row
        _, c_row = self.section.sheet_data.cursor.column

        if top_border:
            if self.from_new_row:
                border_row = r_row
            else:
                border_row = c_row
        else:
            border_row = r_row-1

        return border_row