#coding:utf-8

import re
import xlrd
from xlwt.Style import default_style

from simple_report.core.spreadsheet_section import SpreadsheetSection

class Section(SpreadsheetSection):
    """
    Класс секции отчета в xls
    """

    def __init__(self, sheet, name, begin, end, writer):

        super(Section, self).__init__(sheet, name, begin, end)

        self.sheet = sheet

        self.writer = writer

    def flush(self, params):

        wtrowx = self.sheet.wtrowx

        begin_row, begin_column = self.begin
        end_row, end_column = self.end

        for rdrowx in range(begin_row, end_row + 1):
            for rdcolx in range(begin_column, end_column + 1):
                wtcolx = rdcolx - begin_column

                try:
                    cell = self.writer.rdsheet.cell(rdrowx, rdcolx)
                except IndexError:
                    continue

                val = cell.value

                for key, value in params.items():
                    value = unicode(value)
                    if key in unicode(cell.value):
                        val = val.replace(u'#%s#' % key, value)

                if isinstance(val, basestring):
                    while u'#' in val:
                        val = re.sub(u'#.*#', '', val)

                        if len(val.split('#')) == 2:
                            break

                if wtcolx not in self.writer.wtcols and rdcolx in self.writer.rdsheet.colinfo_map:
                    rdcol = self.writer.rdsheet.colinfo_map[rdcolx]
                    wtcol = self.writer.wtsheet.col(wtcolx)
                    wtcol.width = rdcol.width

                    wtcol.set_style(self.writer.style_list[rdcol.xf_index])
                    wtcol.hidden = rdcol.hidden

                    wtcol.level = rdcol.outline_level
                    wtcol.collapsed = rdcol.collapsed

                    self.writer.wtcols.add(wtcolx)

                cty = self.get_value_type(value=val, default_type=cell.ctype)

                if cty == xlrd.XL_CELL_EMPTY:
                    continue

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

                wtrow = self.writer.wtsheet.row(wtrowx)
                if cty == xlrd.XL_CELL_TEXT:
                    wtrow.set_cell_text(wtcolx, val, style)
                elif cty == xlrd.XL_CELL_NUMBER or cty == xlrd.XL_CELL_DATE:
                    wtrow.set_cell_number(wtcolx, val, style)
                elif cty == xlrd.XL_CELL_BLANK:
                    wtrow.set_cell_blank(wtcolx, style)
                elif cty == xlrd.XL_CELL_BOOLEAN:
                    wtrow.set_cell_boolean(wtcolx, cell.value, style)
                elif cty == xlrd.XL_CELL_ERROR:
                    wtrow.set_cell_error(wtcolx, cell.value, style)
                else:
                    raise Exception
            wtrowx += 1

        self.sheet.wtrowx = wtrowx

    def get_value_type(self, value, default_type=xlrd.XL_CELL_TEXT):
        """
        Возвращаем тип значения для выходного элемента
        """

        try:
            float(value)
            cty = xlrd.XL_CELL_NUMBER
        except ValueError:
            cty = default_type

        return cty
