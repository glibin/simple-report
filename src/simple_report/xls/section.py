#coding:utf-8

from datetime import datetime, date, time

import re
import xlrd
from xlwt.Style import default_style
from simple_report.interface import ISpreadsheetSection

from simple_report.core.exception import XLSReportWriteException
from simple_report.core.spreadsheet_section import SpreadsheetSection, AbstractMerge
from simple_report.xls.cursor import CalculateNextCursorXLS

class Section(SpreadsheetSection, ISpreadsheetSection):
    """
    Класс секции отчета в xls
    """

    def __init__(self, sheet, name, begin, end, writer):

        super(Section, self).__init__(sheet, name, begin, end)

        self.sheet_data = sheet

        self.writer = writer

    def flush(self, params, oriented=ISpreadsheetSection.LEFT_DOWN):

        begin_row, begin_column = self.begin
        end_row, end_column = self.end

        current_col, current_row = self.calc_next_cursor(oriented=oriented)

        for rdrowx in range(begin_row, end_row + 1):
            for rdcolx in range(begin_column, end_column + 1):

                # Вычисляем координаты ячейки для записи.
                wtcolx = current_col + rdcolx - begin_column
                wtrowx = current_row + rdrowx - begin_row

                try:
                    cell = self.writer.rdsheet.cell(rdrowx, rdcolx)
                except IndexError:
                    continue

                val = cell.value

                cty = xlrd.XL_CELL_EMPTY
                for key, value in params.items():
                    if key in unicode(cell.value):
                        # Тип ячейки
                        cty = self.get_value_type(value=value, default_type=cell.ctype)
                        value = unicode(value)
                        val = val.replace(u'#%s#' % key, value)

                if isinstance(val, basestring):
                    while u'#' in val:
                        val = re.sub(u'#.*#', '', val)

                        if len(val.split('#')) == 2:
                            break

                # Копирование всяких свойств из шаблона в результирующий отчет.
                if wtcolx not in self.writer.wtcols and rdcolx in self.writer.rdsheet.colinfo_map:
                    rdcol = self.writer.rdsheet.colinfo_map[rdcolx]
                    wtcol = self.writer.wtsheet.col(wtcolx)
                    wtcol.width = rdcol.width

                    wtcol.set_style(self.writer.style_list[rdcol.xf_index])
                    wtcol.hidden = rdcol.hidden

                    wtcol.level = rdcol.outline_level
                    wtcol.collapsed = rdcol.collapsed

                    self.writer.wtcols.add(wtcolx)

                if cty == xlrd.XL_CELL_EMPTY:
                    continue
                # XF - индексы
                if cell.xf_index is not None:
                    style = self.writer.style_list[cell.xf_index]
                else:
                    style = default_style

                rdcoords2d = rdrowx, rdcolx

                if rdcoords2d in self.writer.merged_cell_top_left_map:

                    rlo, rhi, clo, chi = self.writer.merged_cell_top_left_map[rdcoords2d]
                    assert (rlo, clo) == rdcoords2d
                    self.writer.wtsheet.write_merge(
                        wtrowx, wtrowx + rhi - rlo - 1,
                        wtcolx, wtcolx + chi - clo - 1,
                        val, style)
                    continue

                if rdcoords2d in self.writer.merged_cell_already_set:
                    continue

                self.write_result((wtcolx, wtrowx), val, style, cty)

    def get_width(self):
        """
        """

        begin_row, begin_col = self.begin
        end_row, end_col = self.end

        return end_col - begin_col + 1

    def calc_next_cursor(self, oriented=ISpreadsheetSection.LEFT_DOWN):
        """
        Вычисляем следующее положение курсора.
        """

        begin_row, begin_column = self.begin
        end_row, end_column = self.end

        current_col, current_row = CalculateNextCursorXLS().get_next_cursor(self.sheet_data.cursor, (begin_column, begin_row),
                                                (end_column, end_row), oriented, section=self)

        return current_col, current_row

    #TODO реализовать для поддержки интерфейса ISpreadsheetSection
    def get_all_parameters(self):
        """
        """

    def get_value_type(self, value, default_type=xlrd.XL_CELL_TEXT):
        """
        Возвращаем тип значения для выходного элемента
        """

        cty = default_type

        if isinstance(value, basestring):
            cty = xlrd.XL_CELL_TEXT
        elif isinstance(value, (datetime, date, time)):
            cty = xlrd.XL_CELL_DATE
        elif isinstance(value, bool):
            cty = xlrd.XL_CELL_BOOLEAN
        else:
            cty = xlrd.XL_CELL_NUMBER

        return cty

    def write_result(self, write_coords, value, style, cell_type):
        """
        Выводим в ячейку с координатами write_coords значение value.
        Стиль вывода определяется параметров style
        cty - тип ячейки
        """

        wtcolx, wtrowx = write_coords

        # Вывод
        wtrow = self.writer.wtsheet.row(wtrowx)
        if cell_type == xlrd.XL_CELL_TEXT:
            wtrow.set_cell_text(wtcolx, value, style)
        elif cell_type == xlrd.XL_CELL_NUMBER:
            wtrow.set_cell_number(wtcolx, value, style)
        elif cell_type == xlrd.XL_CELL_DATE:
            wtrow.set_cell_text(wtcolx, value, style)
        elif cell_type == xlrd.XL_CELL_BLANK:
            wtrow.set_cell_blank(wtcolx, style)
        elif cell_type == xlrd.XL_CELL_BOOLEAN:
            wtrow.set_cell_boolean(wtcolx, value, style)
        elif cell_type == xlrd.XL_CELL_ERROR:
            wtrow.set_cell_error(wtcolx, value, style)
        else:
            raise XLSReportWriteException


class MergeXLS(AbstractMerge):
    """
    Конструкция Merge
    """

    def _merge(self):

        self.section.writer.wtsheet.merge(self.begin_row_merge,
                                          self.end_row_merge,
                                          self._begin_merge_col,
                                          self._end_merge_col)

    def _calculate_merge_column(self, column):

        first_section_column = column - self.section.get_width()
        last_section_column = column - 1

        return first_section_column, last_section_column