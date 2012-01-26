# coding: utf-8

import copy
from datetime import datetime
from decimal import Decimal
import locale
import re
from lxml.etree import QName, SubElement
from simple_report.core.shared_table import SharedStringsTable
from simple_report.core.tags import TemplateTags
from simple_report.interface import ISpreadsheetSection
from simple_report.utils import ColumnHelper, get_addr_cell, date_to_float
from simple_report.xlsx.cursor import Cursor

__author__ = 'prefer'


class SheetDataException(Exception):
    """
    """

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

    PREFIX_TAG = '%'

    FIND_PARAMS = re.compile('#\w+#')
    FIND_TEMPLATE_TAGS = re.compile('#{0}\w+{0}#'.format(PREFIX_TAG))

    def __init__(self, sheet_xml, tags, cursor, ns, shared_table):
        # namespace
        self.ns = ns

        # Шаблонные теги
        assert isinstance(tags, TemplateTags)
        self.tags = tags

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


#    def __str__(self):
#        return 'Cursor %s' % self.cursor
#
#    def __repr__(self, ):
#        return self.__str__()

    def flush(self, begin, end, start_cell, params):
        """
        """
        self.set_section(begin, end, start_cell, params)
        self.set_merge_cells(begin, end, start_cell)
        self.set_dimension()

    def set_dimension(self):
        """
        """
        _, row_index = self.cursor.row
        col_index, _ = self.cursor.column

        dimension = 'A1:%s' %\
                    (ColumnHelper.add(col_index, -1) + str(row_index - 1))

        self.write_dimension.set('ref', dimension)

    def _get_merge_cells(self):
        for cell in self.read_merge_cell:
            yield cell.attrib['ref'].split(':')

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

        for begin, end in self._get_merge_cells():

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


    def _find_rows(self, range_rows):
        """

        """

        for i, num_row in enumerate(range_rows):
            row = self.read_data.find(self.XPATH_TEMPLATE_ROW % num_row)
            if not row is None:
                yield i, num_row, row

    def _find_cells(self, range_cols, num_row, row):
        """

        """
        for j, col in enumerate(range_cols):
            cell = row.find(self.XPATH_TEMPLATE_CELL % (col + str(num_row)))
            if not cell is None:
                yield j, col, cell


    def _get_tag_value(self, cell):
        """

        """
        return cell.find(QName(self.ns, 'v'))

    def _get_params(self, cell):
        """

        """
        value = self._get_tag_value(cell)

        if not value is None and cell.get('t') == 's': # 't' = 's' - значит есть значения shared strings
            index_value = int(value.text)
            value_string = self.shared_table.get_value(index_value)

            return self._get_values_by_re(value_string, self.FIND_PARAMS)
        else:
            return []

    def _get_values_by_re(self, value_string, what_found=None):
        """
        """
        if what_found is None:
            # Если значение поиска неопределено выводим поиск для всех параметров
            return self._get_values_by_re(value_string, self.FIND_PARAMS) + \
                    self._get_values_by_re(value_string, self.FIND_TEMPLATE_TAGS)

        who_found_params = what_found.findall(value_string)
        if who_found_params:
            return [found_param for found_param in who_found_params]
        else:
            return []


    def find_all_parameters(self, begin, end):
        """

        """
        range_rows, range_cols = self._range(begin, end)
        for i, num_row, row in self._find_rows(range_rows):
            for j, col, cell in self._find_cells(range_cols, num_row, row):
                for param in self._get_params(cell):
                    yield param


    def _get_tag_formula(self, cell):
        """
        """
        return cell.find(QName(self.ns, 'f'))

    def set_section(self, begin, end, start_cell, params):
        """

        """
        range_rows, range_cols = self._range(begin, end)
        start_column, start_row = start_cell


        for i, num_row, row in self._find_rows(range_rows):

            attrib_row = dict(row.items())

            row_index = str(start_row + i)
            attrib_row['r'] = row_index

            row_el = SubElement(self.write_data, 'row', attrib=attrib_row)

            self.cursor.row = ('A', start_row + i + 1)

            for j, col, cell in self._find_cells(range_cols, num_row, row):

                attrib_cell = dict(cell.items())

                col_index = ColumnHelper.add(start_column, j)

                attrib_cell['r'] = col_index + row_index
                cell_el = SubElement(row_el, 'c', attrib=attrib_cell)

                # Установка курсора
                self.cursor.column = (ColumnHelper.add(col_index, 1), start_row)

                # Перенос формул
                formula = self._get_tag_formula(cell)
                if formula is not None:
                    formula_el = SubElement(cell_el, 'f')
                    formula_el.text = formula.text

                    # Если есть формула, то значение является вычисляемым параметром и не сильно интересует
                    continue

                value = self._get_tag_value(cell)
                if not value is  None:
                    value_el = SubElement(cell_el, 'v')

                    if attrib_cell.get('t') in ('n', None): # number

                        value_el.text = value.text

                    elif attrib_cell.get('t') == 's': # 't' = 's' - значит есть значения shared strings

                        index_value = int(value.text)
                        value_string = self.shared_table.get_value(index_value)

                        who_found_params = self._get_values_by_re(value_string)

                        is_int = False
                        if who_found_params:
                            for found_param in who_found_params:
                                param_name = found_param[1:-1]

                                param_value = params.get(param_name)

                                # Находим теги шаблонов, если есть таковые
                                if param_name[0] == self.PREFIX_TAG and param_name[-1] == self.PREFIX_TAG:
                                    param_value = self.tags.get(param_name[1:-1])

                                if isinstance(param_value, datetime) and found_param == value_string:

                                    # В OpenXML хранится дата относительно 1900 года
                                    days = date_to_float(param_value)
                                    if days > 0:
                                        # Дата конвертируется в int, начиная с 31.12.1899
                                        is_int = True
                                        cell_el.attrib['t'] = 'n' # type - number
                                        value_el.text = unicode(days)
                                    else:
                                        date_less_1900 = '%s.%s.%s' % (param_value.date().day,
                                                                       param_value.date().month, param_value.date().year,)
                                        # strftime(param_value, locale.nl_langinfo(locale.D_FMT)) - неработает для 1900 и ниже
                                        value_string = value_string.replace(found_param, unicode(date_less_1900))


                                elif isinstance(param_value, (int, float, Decimal)) and found_param == value_string:
                                    # В первую очередь добавляем числовые значения
                                    is_int = True

                                    cell_el.attrib['t'] = 'n' # type - number
                                    value_el.text = unicode(param_value)

                                elif param_value:
                                    # Строковые параметры

                                    value_string = value_string.replace(found_param, unicode(param_value))

                                else:
                                    # Не передано значение параметра
                                    value_string = value_string.replace(found_param, '')

                            if not is_int:
                                # Добавим данные в shared strings

                                new_index = self.shared_table.get_new_index(index_value)
                                value_el.text = new_index
                                self.shared_table.new_elements_list[int(new_index)] = value_string


                        else:
                            # Параметры в поле не найдены

                            index = self.shared_table.get_new_index(value.text)
                            value_el.text = index

                    elif attrib_cell.get('t'):
                        raise SheetDataException("Unknown value '%s' for tag t" % attrib_cell.get('t'))

    def _range(self, begin, end):
        """
        """

        # Если есть объединенная ячейка, и она попадает на конец секции, то адресс конца секции записывается как начало
        # объединенной ячейки
        for begin_merge, end_merge in self._get_merge_cells():
            addr = get_addr_cell(begin_merge)
            if addr == end:
                end = get_addr_cell(end_merge)
                break

        rows = begin[1], end[1] + 1
        cols = begin[0], end[0]

        range_rows = xrange(*rows)
        range_cols = list(ColumnHelper.get_range(*cols))

        return range_rows, range_cols


    def new_sheet(self):
        """
        """
        return self._write_xml



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
        return self.sheet_data.find_all_parameters(self.begin, self.end)
