#coding: utf-8
__author__ = 'prefer'

from simple_report.core.cursor import AbstractCursor, AbstractCalculateNextCursor
from simple_report.utils import ColumnHelper
from simple_report.core.exception import XLSXReportWriteException

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


class MarkerPosition(object):
    """
    Указывает на левую верхнюю ячейку последней выведенной секции
    """

    def __init__(self):

        self.column = 'A'
        self.row = 1


class DirectionSet(object):
    """
    Множество направлений вывода
    """

    # N  - север
    # NE - северо-восток
    # E  - восток
    # SE - юго-восток
    # S  - юг
    # SW - юго-запад
    # W  - запад
    # NW - северо-запад
    # LD - первая колонка следующей строки
    # RU - первая колонка следующего столбца
    N, NE, E, SE, S, SW, W, NW, LD, RU = range(1, 11)


class CalculateNextMarker(object):
    """
    """

    @staticmethod
    def get_next_marker(marker, direction, current_section_size, prev_section_size):
        """
        """

        width, height = current_section_size
        last_width, last_height = prev_section_size

        if direction == DirectionSet.N:
            if marker.row < height:
                raise XLSXReportWriteException
            marker.row -= height
        elif direction == DirectionSet.S:
            marker.row += last_height
        elif direction == DirectionSet.W:
            if marker.column < width:
                raise XLSXReportWriteException
            marker.column = ColumnHelper.column_to_number(ColumnHelper.column_to_number(marker.column) - width)
        elif direction == DirectionSet.E:
            marker.column = ColumnHelper.add(marker.column, last_width)
        elif direction == DirectionSet.NE:
            if marker.row < height:
                raise XLSXReportWriteException
            marker.row -= height
            marker.column = ColumnHelper.add(marker.column, last_width)
        elif direction == DirectionSet.SE:
            marker.row += last_height
            marker.column = ColumnHelper.add(marker.column, last_width)
        elif direction == DirectionSet.SW:
            if marker.column < width:
                raise XLSXReportWriteException
            marker.row += last_height
            marker.column = ColumnHelper.number_to_column(ColumnHelper.column_to_number(marker.column) - width)
        elif direction == DirectionSet.NW:
            if marker.column < width or marker.row < height:
                raise XLSXReportWriteException
            marker.row -= height
            marker.column = ColumnHelper.number_to_column(ColumnHelper.column_to_number(marker.column) - width)
        elif direction == DirectionSet.LD:
            marker.row += last_height
            marker.column = 'A'
        elif direction == DirectionSet.RU:
            marker.row = 1
            marker.column = ColumnHelper.add(marker.column, last_width)


class CalculateNextCursor(AbstractCalculateNextCursor):
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