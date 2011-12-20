# coding: utf-8

import copy
from lxml.etree import QName, SubElement
from simple_report.core.shared_table import SharedStringsTable
from simple_report.interface import ISpreadsheetSection
from simple_report.utils import ColumnHelper, get_addr_cell
from simple_report.xlsx.cursor import Cursor

__author__ = 'prefer'


class SheetData(object):
    u"""
    self.read_data:
        <sheetData>
            <row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.8" outlineLevel="0" r="5">
                <c r="C5" s="1" t="s">
                    <v>0</v>
                </c>
                <c r="D5" s="1"/>
                <c r="E5" s="1"/>
            </row>
            <row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.8" outlineLevel="0" r="6">
                <c r="C6" s="0" t="s">
                    <v>1</v>
                </c>
            </row>
        </sheetData>

    self.read_dimension:
        <dimension ref="B5:I10"/>

    self.read_merge_cell:
        <mergeCells count="1">
            <mergeCell ref="C5:E5"/>
        </mergeCells>

    Данные для подобных write атрибутов должны быть такой же структуры
    """

    XPATH_TEMPLATE_ROW = '*[@r="%d"]'
    XPATH_TEMPLATE_CELL = '*[@r="%s"]'

    def __init__(self, sheet_xml, cursor, ns, shared_table):
        # namespace
        self.ns = ns

        assert isinstance(cursor, Cursor)
        self.cursor = cursor

        assert isinstance(shared_table, SharedStringsTable)
        self.shared_table = shared_table

        self._read_xml = sheet_xml

        self.read_data = sheet_xml.find(QName(self.ns, 'sheetData'))
        self.read_dimension = sheet_xml.find(QName(self.ns, 'dimension'))
        self.read_merge_cell = sheet_xml.find(QName(self.ns, 'mergeCells'))

        self._write_xml = copy.deepcopy(sheet_xml)

        # Ссылка на тег данных строк и столбцов листа с очищенными значениями
        self.write_data = self._write_xml.find(QName(self.ns, 'sheetData'))
        if not self.write_data is None:
            self.write_data.clear()

        # Ссылка на размеры листа
        self.write_dimension = self._write_xml.find(QName(self.ns, 'dimension'))

        # Ссылка на объединенные ячейки листа с очищенными значениями
        self.write_merge_cell = self._write_xml.find(QName(self.ns, 'mergeCells'))
        if not self.write_merge_cell is None:
            self.write_merge_cell.clear()


    def __str__(self):
        return 'Cursor %s' % self.cursor

    def __repr__(self, ):
        return self.__str__()

    def flush(self, begin, end, start_cell, params):
        """
        """
        indexes = self.set_section(begin, end, start_cell)
        self.set_merge_cells(begin, end, start_cell)
        self.set_dimension()

        self.set_params(indexes, params)

    def set_dimension(self):
        """
        """
        _, row_index = self.cursor.row
        col_index, _ = self.cursor.column

        dimension = 'A1:%s' %\
                    (ColumnHelper.add(col_index, -1) + str(row_index - 1))

        self.write_dimension.set('ref', dimension)

    def set_merge_cells(self, section_begin, section_end, start_cell):
        """
        """

        def cell_dimensions(section, merge_cell, start_cell):
            """
            """

            section_begin_col, section_begin_row = section

            start_col, start_row = start_cell

            begin_col, begin_row = merge_cell

            new_begin_row = start_row + begin_row - section_begin_row
            new_begin_col = ColumnHelper.add(start_col, ColumnHelper.difference(begin_col, section_begin_col))

            return new_begin_col + str(new_begin_row)

        range_rows, range_cols = self._range(section_begin, section_end)

        for cell in self.read_merge_cell:
            begin, end = cell.attrib['ref'].split(':')

            begin_col, begin_row = get_addr_cell(begin)
            end_col, end_row = get_addr_cell(end)

            # Если объединяемый диапазон лежит внутри секции
            if (begin_col in range_cols and end_col in range_cols and
                begin_row in range_rows and end_row in range_rows):
                begin_merge = cell_dimensions(section_begin, (begin_col, begin_row), start_cell)
                end_merge = cell_dimensions(section_begin, (end_col, end_row), start_cell)

                attrib = {'ref': ':'.join((begin_merge, end_merge))}
                SubElement(self.write_merge_cell, 'mergeCell', attrib)

        count_merge_cells = len(self.write_merge_cell)
        if count_merge_cells:
            self.write_merge_cell.set('count', str(count_merge_cells))

    def set_section(self, begin, end, start_cell):
        """
        @param begin: Начала секции
        @param end: Конец секции
        @param start_cell: Начало вывода секции

        """
        # TODO: Отрефакторить и разбить на несколько методов. Так же приспособить для поиска параметров
        list_index = []

        range_rows, range_cols = self._range(begin, end)

        start_column, start_row = start_cell

        for i, num_row in enumerate(range_rows):
            row = self.read_data.find(self.XPATH_TEMPLATE_ROW % num_row)
            if not row is None:
                # Только если есть строки
                attrib_row = dict(row.items())

                row_index = str(start_row + i)
                attrib_row['r'] = row_index

                row_el = SubElement(self.write_data, 'row', attrib=attrib_row)

                self.cursor.row = ('A', start_row + i + 1)

                for j, col in enumerate(range_cols):
                    # Столбцы начинают выводится каждый раз от начала курсора

                    cell = row.find(self.XPATH_TEMPLATE_CELL % (col + str(num_row)))
                    if not cell is None:
                    # Только если есть ячейки
                        attrib_cell = dict(cell.items())

                        col_index = ColumnHelper.add(start_column, j)

                        attrib_cell['r'] = col_index + row_index
                        cell_el = SubElement(row_el, 'c', attrib=attrib_cell)

                        # Установка курсора
                        self.cursor.column = (ColumnHelper.add(col_index, 1), start_row)

                        value = cell.find(QName(self.ns, 'v'))
                        if not value is None:
                            # Только если есть значения в ячейках

                            value_el = SubElement(cell_el, 'v')
                            index = self.shared_table.get_new_index(value.text)
                            value_el.text = index

                            list_index.append(int(index))

        return list_index


    def _range(self, begin, end):
        """
        """

        rows = begin[1], end[1] + 1
        cols = begin[0], end[0]

        range_rows = xrange(*rows)
        range_cols = list(ColumnHelper.get_range(*cols))

        return range_rows, range_cols


    def new_sheet(self):
        """
        """
        return self._write_xml

    def set_params(self, indexes, params):
        """
        """
        self.shared_table.set_params(indexes, params)



class Section(ISpreadsheetSection):
    """
    """

    def __init__(self, sheet_data, name, begin=None, end=None):
        """
        @param sheet_data: Данные листа
        @param name: Название секции
        @param begin: Начало секции, пример ('A', 1)
        @param end: Конец секции, пример ('E', 6)
        """
        assert begin or end

        self.name = name
        self.begin = begin
        self.end = end

        # Ссылка на курсор листа. Метод flush вставляет данные относительно курсора
        # и меняет его местоположение
        assert isinstance(sheet_data, SheetData)
        self.sheet_data = sheet_data

    def __str__(self):
        return 'Section "{0} - ({1},{2}) \n\t sheet_data - {3}" '.format(
            self.name, self.begin, self.end, self.sheet_data)

    def __repr__(self):
        return self.__str__()


    def flush(self, params, oriented=ISpreadsheetSection.VERTICAL):
        """
        """
        assert isinstance(params, dict)
        assert oriented in (Section.VERTICAL, Section.GORIZONTAL)

        # Тут смещение курсора, копирование данных из листа и общих строк
        # Генерация новых данных и новых общих строк

        start_cell = self.sheet_data.cursor.row if oriented == Section.VERTICAL else self.sheet_data.cursor.column

        self.sheet_data.flush(self.begin, self.end, start_cell, params)

    def get_all_parameters(self):
        u"""
        Возвращает все параметры секции
        """

