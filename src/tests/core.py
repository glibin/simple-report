#coding: utf-8
from datetime import datetime
import sys
sys.path.append('/home/vahotin/dev/simple_report/src')
from simple_report.converter.abstract import FileConverter
from simple_report.core.tags import TemplateTags
from simple_report.interface import ISpreadsheetSection
from simple_report.xlsx.section import Section
from simple_report.xlsx.spreadsheet_ml import SectionException, SectionNotFoundException

sys.path.append('.')

from test_oo_wrapper import TestOO
from test_utils import skip_python26
from test_pko import TestPKO
from oborot import OperationsJournalReportFactory
from test_pagebreaks import TestPagebreaks

sys.path.append('../')

import os
import unittest

from simple_report.report import (SpreadsheetReport, ReportGeneratorException, DocumentReport,
                                  )
from simple_report.xls.document import DocumentXLS
from simple_report.converter.abstract import FileConverter
from simple_report.utils import ColumnHelper, date_to_float


class TestXLSX(object):
    """

    """

    # Разные директории с разными файлами под linux и под windows
    SUBDIR = None

    def setUp(self):
        assert self.SUBDIR
        self.src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data', self.SUBDIR, 'xlsx', )
        self.dst_dir = self.src_dir

        self.test_files = dict([(path, os.path.join(self.src_dir, path))
        for path in os.listdir(self.src_dir) if path.startswith('test')])


    @skip_python26
    def test_range_cols(self):
        section_range = list(ColumnHelper.get_range(('ALC'), ('AVB')))
        self.assertIn('ALC', section_range)
        self.assertIn('AVB', section_range)
        self.assertIn('AVA', section_range)
        self.assertNotIn('ALA', section_range)
        self.assertNotIn('AVC', section_range)

        section_range = list(ColumnHelper.get_range(('X'), ('AB')))
        self.assertIn('X', section_range)
        self.assertIn('Y', section_range)
        self.assertIn('AA', section_range)
        self.assertNotIn('Q', section_range)
        self.assertNotIn('AC', section_range)
        self.assertEqual(len(section_range), 5)

        section_range = list(ColumnHelper.get_range(('B'), ('CBD')))
        self.assertIn('B', section_range)
        self.assertIn('C', section_range)
        self.assertIn('AAA', section_range)
        self.assertIn('ZZ', section_range)
        self.assertIn('ABA', section_range)
        self.assertIn('CBD', section_range)
        self.assertIn('CAA', section_range)
        self.assertIn('BZZ', section_range)
        self.assertNotIn('CBE', section_range)

        section_range = list(ColumnHelper.get_range(('BCCC'), ('BCCD')))
        self.assertIn('BCCC', section_range)
        self.assertIn('BCCD', section_range)
        self.assertNotIn('A', section_range)
        self.assertNotIn('BCCE', section_range)
        self.assertEqual(len(section_range), 2)

    @skip_python26
    def test_workbook(self):
        src = self.test_files['test-simple.xlsx']
        dst = os.path.join(self.dst_dir, 'res-simple.xlsx')
        if os.path.exists(dst):
            os.remove(dst)
        self.assertEqual(os.path.exists(dst), False)

        report = SpreadsheetReport(src, tags=TemplateTags(test_tag=222))

        #self.assertGreater(len(report._wrapper.sheets), 0)

        self.assertNotEqual(report._wrapper.workbook, None)
        self.assertNotEqual(report._wrapper.workbook.shared_strings, None)

        # Тестирование получения секции
        section_a1 = report.get_section('A1')
        self.assertIsInstance(section_a1, Section)

        with self.assertRaises(SectionNotFoundException):
            report.get_section('G1')

        section_a1.flush({'user': u'Иванов Иван',
                          'date_now': 1})

        s_gor = report.get_section('GOR')
        s_gor.flush({'col': u'Данные'}, oriented=s_gor.HORIZONTAL)

        for i in range(10):
            report.get_section('B1').flush({'nbr': i,
                                            'fio': u'Иванов %d' % i,
                                            'sector': u'Какой-то сектор'})

            s_gor_str = report.get_section('GorStr')
            s_gor_str.flush({'g': i + i}, oriented=s_gor.HORIZONTAL)
            s_gor_str.flush({'g': i * i}, oriented=s_gor.HORIZONTAL)

        report.get_section('C1').flush({'user': u'Иван'})

        with self.assertRaises(ReportGeneratorException):
            report.build(dst, FileConverter.XLS)

        report.build(dst)

        self.assertEqual(os.path.exists(dst), True)


    def test_workbook_with_2_6_python(self):
        src = self.test_files['test-simple.xlsx']
        dst = os.path.join(self.dst_dir, 'res-simple.xlsx')
        if os.path.exists(dst):
            os.remove(dst)
        self.assertEqual(os.path.exists(dst), False)

        report = SpreadsheetReport(src, tags=TemplateTags(test_tag=11))

        section_a1 = report.get_section('A1')
        section_a1.flush({'user': u'Иванов Иван',
                          'date_now': 1})

        s_gor = report.get_section('GOR')
        s_gor.flush({'col': u'Данные'}, oriented=s_gor.HORIZONTAL)

        for i in range(10):
            report.get_section('B1').flush({'nbr': i,
                                            'fio': u'Иванов %d' % i,
                                            'sector': u'Какой-то сектор'})

            s_gor_str = report.get_section('GorStr')
            s_gor_str.flush({'g': i + i}, oriented=s_gor.HORIZONTAL)
            s_gor_str.flush({'g': i * i}, oriented=s_gor.HORIZONTAL)

        report.get_section('C1').flush({'user': u'Иван'})
        report.build(dst)
        self.assertEqual(os.path.exists(dst), True)

class TestLinuxXLSX(TestXLSX, TestOO, TestPKO, TestPagebreaks, unittest.TestCase):
    SUBDIR = 'linux'

    @skip_python26
    def test_fake_section(self):
        src = self.test_files['test-simple-fake-section.xlsx']
        with self.assertRaises(SectionException):
            report = SpreadsheetReport(src)
            report.build(src)

    def test_merge_cells(self):
        src = self.test_files['test-merge-cells.xlsx']
        dst = os.path.join(self.dst_dir, 'res-merge-cells.xlsx')

        report = SpreadsheetReport(src)

        report.get_section('head').flush({'kassa_za': u'Ноябрь'})

        for i in range(10):
            report.get_section('table_dyn').flush({'doc_num': i})

        report.get_section('foot').flush({'glavbuh': u'Иван'})

        report.build(dst)


    def test_383_value(self):
        src = self.test_files['test-383.xlsx']
        dst = os.path.join(self.dst_dir, 'res-383.xlsx')

        report = SpreadsheetReport(src)

        report.get_section('header').flush({'period': u'Ноябрь'})

        for i in range(10):
            report.get_section('row').flush({'begin_year_debet': -i})

        report.get_section('footer').flush({'glavbuh': u'Иван'})

        report.build(dst)

    @skip_python26
    def test_main_parameters(self):
        src = self.test_files['test-main_book.xlsx']
        dst = os.path.join(self.dst_dir, 'res-main_book.xlsx')

        report = SpreadsheetReport(src)

        params_header = list(report.get_section('header').get_all_parameters())
        self.assertEqual(0, len(params_header))

        params_row = list(report.get_section('row').get_all_parameters())
        self.assertEqual(13, len(params_row))
        self.assertIn(u'#num#', params_row)
        self.assertIn(u'#account_name#', params_row)
        self.assertIn(u'#journal_num#', params_row)

        params_footer = list(report.get_section('footer').get_all_parameters())
        self.assertEqual(12, len(params_footer))
        self.assertIn(u'#begin_year_debet_sum#', params_footer)
        self.assertIn(u'#glavbuh#', params_footer)
        self.assertIn(u'#username#', params_footer)


    def test_empty_cell(self):
        """

        """
        file_name = 'test-empty-section.xlsx'
        template_name = self.test_files[file_name]
        report = SpreadsheetReport(template_name)

        header = report.get_section(u'row')

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 2))
        self.assertEqual(header.sheet_data.cursor.column, ('B', 1))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 2))
        self.assertEqual(header.sheet_data.cursor.column, ('C', 1))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 2))
        self.assertEqual(header.sheet_data.cursor.column, ('D', 1))

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 3))
        self.assertEqual(header.sheet_data.cursor.column, ('B', 2))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 3))
        self.assertEqual(header.sheet_data.cursor.column, ('C', 2))

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 4))
        self.assertEqual(header.sheet_data.cursor.column, ('B', 3))

        #result_file, result_url = create_office_template_tempnames(template_name)
        res_file_name = 'res-' + file_name
        dst = os.path.join(self.dst_dir, res_file_name)
        report.build(dst)


    def test_wide_cell_1(self):
        """

        """
        file_name = 'test-wide-section-1.xlsx'
        template_name = self.test_files[file_name]
        report = SpreadsheetReport(template_name)

        header = report.get_section(u'row')

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 2))
        self.assertEqual(header.sheet_data.cursor.column, ('B', 1))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 2))
        self.assertEqual(header.sheet_data.cursor.column, ('C', 1))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 2))
        self.assertEqual(header.sheet_data.cursor.column, ('D', 1))

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 3))
        self.assertEqual(header.sheet_data.cursor.column, ('B', 2))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 3))
        self.assertEqual(header.sheet_data.cursor.column, ('C', 2))

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 4))
        self.assertEqual(header.sheet_data.cursor.column, ('B', 3))

        #result_file, result_url = create_office_template_tempnames(template_name)
        res_file_name = 'res-' + file_name
        dst = os.path.join(self.dst_dir, res_file_name)
        report.build(dst)


    def test_wide_cell_2(self):
        file_name = 'test-wide-section-2.xlsx'
        template_name = self.test_files[file_name]
        report = SpreadsheetReport(template_name)

        header = report.get_section(u'row')

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 4))
        self.assertEqual(header.sheet_data.cursor.column, ('D', 1))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 4))
        self.assertEqual(header.sheet_data.cursor.column, ('G', 1))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 4))
        self.assertEqual(header.sheet_data.cursor.column, ('J', 1))

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 7))
        self.assertEqual(header.sheet_data.cursor.column, ('D', 4))

        header.flush({}, oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(header.sheet_data.cursor.row, ('A', 7))
        self.assertEqual(header.sheet_data.cursor.column, ('G', 4))

        header.flush({})
        self.assertEqual(header.sheet_data.cursor.row, ('A', 10))
        self.assertEqual(header.sheet_data.cursor.column, ('D', 7))

        #result_file, result_url = create_office_template_tempnames(template_name)
        res_file_name = 'res-' + file_name
        dst = os.path.join(self.dst_dir, res_file_name)
        report.build(dst)

        return res_file_name

    def test_without_merge_cells(self):
        file_name = 'test-main_template.xlsx'
        template_name = self.test_files[file_name]
        report = SpreadsheetReport(template_name)

        head_section = report.get_section('head')
        for head in range(10):
            head_section.flush({'head_name': str(head)}, 1)

        #result_file, result_url = create_office_template_tempnames(template_name)
        res_file_name = 'res-' + file_name
        dst = os.path.join(self.dst_dir, res_file_name)
        report.build(dst)

        return res_file_name


    def test_purchases_book(self):
        """

        """

        enterprise_name = u'Мегаэнтерпрайз'
        inn = 123123123
        kpp = 123123
        date_start = datetime.now()
        date_end = datetime.now()

        file_name = 'test-purchases_book.xlsx'
        template_name = self.test_files[file_name]
        report = SpreadsheetReport(template_name)

        header = report.get_section('header')
        header.flush({
            'enterprise_name': enterprise_name,
            'inn': inn,
            'kpp': kpp,
            'date_start': date_start,
            'date_end': date_end
        })

        res_file_name = 'res-' + file_name
        dst = os.path.join(self.dst_dir, res_file_name)
        report.build(dst)


    def test_operations_journal(self):
        """

        """
        template_name = "test-operations_journal.xlsx"
        path = self.test_files[template_name]

        report = OperationsJournalReportFactory(path).generate()

        res_file_name = 'res-' + template_name
        dst = os.path.join(self.dst_dir, res_file_name)

        return report.build(dst)

    def test_cursor(self):
        """
        Правильность вычисления курсора
        """

        src = self.test_files['test-simple.xlsx']
        dst = os.path.join(self.dst_dir, 'res-simple.xlsx')
        if os.path.exists(dst):
            os.remove(dst)
        self.assertEqual(os.path.exists(dst), False)

        report = SpreadsheetReport(src, tags=TemplateTags(test_tag=222))

        section_a1 = report.get_section('A1')
        section_a1.flush({'user': u'Иванов Иван',
                          'date_now': 1})

        # Проверяем курсор для колонки
        self.assertEqual(section_a1.sheet_data.cursor.row, ('A', 5))
        # Проверяем курсор для строки
        self.assertEqual(section_a1.sheet_data.cursor.column, ('D', 1))

        s_gor = report.get_section('GOR')
        s_gor.flush({'col': u'Данные'}, oriented=s_gor.HORIZONTAL)

        # Проверяем курсор для колонки
        self.assertEqual(s_gor.sheet_data.cursor.row, ('A', 5))
        # Проверяем курсор для строки
        self.assertEqual(s_gor.sheet_data.cursor.column, ('E', 1))

        for i in range(10):
            report.get_section('B1').flush({'nbr': i,
                                            'fio': u'Иванов %d' %i,
                                            'sectior': u'Какой-то сектор'})

            self.assertEqual(s_gor.sheet_data.cursor.row, ('A', i + 6))
            self.assertEqual(s_gor.sheet_data.cursor.column, ('D', i + 5))

            s_gor_str = report.get_section('GorStr')
            s_gor_str.flush({'g': i + 1}, oriented=s_gor.HORIZONTAL)

            self.assertEqual(s_gor.sheet_data.cursor.row, ('A', i + 6))
            self.assertEqual(s_gor.sheet_data.cursor.column, ('E', i + 5))

            s_gor_str.flush({'g': i * i}, oriented=s_gor.HORIZONTAL)

            self.assertEqual(s_gor.sheet_data.cursor.row, ('A', i + 6))
            self.assertEqual(s_gor.sheet_data.cursor.column, ('F', i + 5))

        section_last = report.get_section('C1')
        section_last.flush({'user': u'Иван'})
        self.assertEqual(section_last.sheet_data.cursor.row, ('A', 16))
        self.assertEqual(section_last.sheet_data.cursor.column, ('D', 15))

class TestWriteXLSX(unittest.TestCase):
    """
    Тестируем правильность вывода для XLSX
    """

    SUBDIR = 'linux'

    def setUp(self):
        assert self.SUBDIR
        self.src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data', self.SUBDIR, 'xlsx', )
        self.dst_dir = self.src_dir

        self.test_files = dict([(path, os.path.join(self.src_dir, path))
        for path in os.listdir(self.src_dir) if path.startswith('test')])

    def test_left_down(self, report=None):
        if report is None:
            return
        for i in range(2):
            section1 = report.get_section('Section1')
            section1.flush({'section1': i}, oriented=Section.LEFT_DOWN)
            self.assertEqual(section1.sheet_data.cursor.row, ('A', 2*i + 3))
            self.assertEqual(section1.sheet_data.cursor.column, ('C', 2*i + 1))

    def test_left_down2(self, report=None):
        if report is None:
            return
        for i in range(2):
            section3 = report.get_section('Section3')
            section3.flush({'section3': 100}, oriented=Section.LEFT_DOWN)
            self.assertEqual(section3.sheet_data.cursor.row, ('A', 2*i + 11))
            self.assertEqual(section3.sheet_data.cursor.column, ('C', 2*i + 9))

    def test_right_up(self, report=None):
        if report is None:
            return
        section1 = report.get_section('Section1')
        section1.flush({'section1': 2}, oriented=Section.RIGHT_UP)
        self.assertEqual(section1.sheet_data.cursor.row, ('C', 3))
        self.assertEqual(section1.sheet_data.cursor.column, ('E', 1))

    def test_vertical(self, report=None):
        if report is None:
            return
        for i in range(3):
            section2 = report.get_section('Section2')
            section2.flush({'section2': i}, oriented=Section.VERTICAL)
            self.assertEqual(section2.sheet_data.cursor.row, ('C', 2*i + 5))
            self.assertEqual(section2.sheet_data.cursor.column, ('E', 2*i + 3))

    def test_horizontal(self, report=None):
        if report is None:
            return
        for i in range(3):
            section3 = report.get_section('Section3')
            section3.flush({'section3': i}, oriented=Section.HORIZONTAL)
            self.assertEqual(section3.sheet_data.cursor.row, ('C', 9))
            self.assertEqual(section3.sheet_data.cursor.column,
                (ColumnHelper.add('G', 2*i), 7))

    def test_report_write(self):

        src = self.test_files['test-report-output.xlsx']
        dst = os.path.join(self.dst_dir, 'res-report-output.xlsx')
        if os.path.exists(dst):
            os.remove(dst)
        self.assertEqual(os.path.exists(dst), False)

        report = SpreadsheetReport(src)
        self.test_left_down(report)
        self.test_right_up(report)
        self.test_vertical(report)
        self.test_horizontal(report)
        self.test_left_down2(report)

        return report.build(dst)

class TestWriteXLS(unittest.TestCase):
    """
    Тестируем правильность вывода для XSL
    """

    SUBDIR = 'linux'

    def setUp(self):
        assert self.SUBDIR
        self.src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data', self.SUBDIR, 'xls', )
        self.dst_dir = self.src_dir

        self.test_files = dict([(path, os.path.join(self.src_dir, path))
        for path in os.listdir(self.src_dir) if path.startswith('test')])

    def test_left_down(self, report=None):
        if report is None:
            return
        for i in range(2):
            section1 = report.get_section('Section1')
            section1.flush({'section1': i}, oriented=Section.LEFT_DOWN)
            self.assertEqual(section1.sheet.cursor.row, (0, 2*i + 2))
            self.assertEqual(section1.sheet.cursor.column, (2, 2*i))

    def test_left_down2(self, report=None):
        if report is None:
            return
        for i in range(2):
            section3 = report.get_section('Section3')
            section3.flush({'section3': 100}, oriented=Section.LEFT_DOWN)
            self.assertEqual(section3.sheet.cursor.row, (0, 2*i + 10))
            self.assertEqual(section3.sheet.cursor.column, (2, 2*i + 8))

    def test_right_up(self, report=None):
        if report is None:
            return
        section1 = report.get_section('Section1')
        section1.flush({'section1': 2}, oriented=Section.RIGHT_UP)
        self.assertEqual(section1.sheet.cursor.row, (2, 2))
        self.assertEqual(section1.sheet.cursor.column, (4, 1))

    def test_vertical(self, report=None):
        if report is None:
            return
        for i in range(3):
            section2 = report.get_section('Section2')
            section2.flush({'section2': i}, oriented=Section.VERTICAL)
            self.assertEqual(section2.sheet.cursor.row, (2, 2*(i+1) + 2))
            self.assertEqual(section2.sheet.cursor.column, (4, 2*(i+1)))

    def test_horizontal(self, report=None):
        if report is None:
            return
        for i in range(3):
            section3 = report.get_section('Section3')
            section3.flush({'section3': i}, oriented=Section.HORIZONTAL)
            self.assertEqual(section3.sheet.cursor.row, (2, 8))
            self.assertEqual(section3.sheet.cursor.column,
                (6 + 2*i, 6))

    def test_report_write(self):

        src = self.test_files['test-report-output.xls']
        dst = os.path.join(self.dst_dir, 'res-report-output.xls')
        if os.path.exists(dst):
            os.remove(dst)

        report = SpreadsheetReport(src, wrapper=DocumentXLS, type=FileConverter.XLS)
        self.test_left_down(report)
        self.test_right_up(report)
        self.test_vertical(report)
        self.test_horizontal(report)
        self.test_left_down2(report)

        return report.build(dst)

class TestReportFormatXLS(unittest.TestCase):
    """
    Тест на работоспособность отчета формата XLS
    """

    SUBDIR = 'linux'

    def setUp(self):
        assert self.SUBDIR
        self.src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data', self.SUBDIR, 'xls', )
        self.dst_dir = self.src_dir

        self.test_files = dict([(path, os.path.join(self.src_dir, path))
        for path in os.listdir(self.src_dir) if path.startswith('test')])

    def test_spreadsheet_with_flag(self):
        """
        Тест на использование класса SpreadsheetReport с переданным в конструктор wrapper-ом.
        """

        src = self.test_files['test_xls.xls']
        dst = os.path.join(self.dst_dir, 'res-test_xls.xls')
        if os.path.exists(dst):
            os.remove(dst)

        report = SpreadsheetReport(src, wrapper=DocumentXLS, type=FileConverter.XLS)

        section1 = report.get_section('Section1')
        section1.flush({'tag1': 1})

        report.workbook.active_sheet = 1

        section2 = report.get_section('Section2')
        for i in range(10):
            section2.flush({'tag2': i})

        for i in range(10):
            section2.flush({'tag2': str(10)}, oriented=Section.HORIZONTAL)

        return report.build(dst)


class TestWindowsXLSX(TestXLSX, unittest.TestCase):
    SUBDIR = 'win'


class TestLinuxDOCX(unittest.TestCase):
    """

    """

    def setUp(self):
        self.src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data', 'linux', 'docx', )
        self.dst_dir = self.src_dir

        self.test_files = dict([(path, os.path.join(self.src_dir, path))
        for path in os.listdir(self.src_dir) if path.startswith('test')])


    def test_simple_docx(self):
        """
        """

        template_name = 'test-sluzh.docx'
        path = self.test_files[template_name]
        doc = DocumentReport(path)

        res_file_name = 'res-' + template_name
        dst = os.path.join(self.dst_dir, res_file_name)

        doc.build(dst, {'Employee_name': u'Иванов И.И.', 'region_name': u'Казань'})
        self.assertEqual(os.path.exists(dst), True)

    def test_spreadsheet_docx(self):
        """
        Текст внутри таблицы
        """

        template_name = 'test_spreadsheet.docx'
        path = self.test_files[template_name]
        doc = DocumentReport(path)

        res_file_name = 'res-' + template_name
        dst = os.path.join(self.dst_dir, res_file_name)

        tag1 = doc.get_all_parameters().next()
        self.assertEqual(tag1, '#sometext#')

        doc.build(dst, {'sometext': u'Некий текст'})
        self.assertEqual(os.path.exists(dst), True)

    def test_picture_docx(self):
        """
        Текст внутри прямоугольника
        """

        template_name = 'test_rect.docx'
        path = self.test_files[template_name]
        doc = DocumentReport(path)

        res_file_name = 'res-' + template_name
        dst = os.path.join(self.dst_dir, res_file_name)

        tags = []
        for tag in doc.get_all_parameters():
            tags.append(tag)

        self.assertFalse(tags[0] != '#brandgroupname#' and tags[0] != '#category#')
        self.assertFalse(tags[1] != '#brandgroupname#' and tags[1] != '#category#')

        doc.build(dst, {'brandgroupname': u'Брэнд', 'category': u'Категория'})
        self.assertEqual(os.path.exists(dst), True)

    def test_picture_shape(self):

        template_name = 'test_pict_shape_2.docx'
        path = self.test_files[template_name]

        res_file_name = 'res-pict_shape.docx'
        dst = os.path.join(self.dst_dir, res_file_name)

        report = DocumentReport(path)
        params = {}

        params['fname'] = '1'
        params['sname'] = '2'
        params['pname'] = '3'
        params['issued_by'] = '4'
        params['date_of_birth'] = '5'

        params['date_start_day'] = '6'
        params['date_start_month'] = '7'
        params['date_start_year'] = '8'
        params['date_start'] = '9'
        params['date_end_day'] = '10'
        params['date_end_month'] = '11'
        params['date_end_year'] = '12'
        params['date_end'] = '13'
        params['region_number'] = '14'
        params['date_start_plus'] = '15'
        params['date_start_plus_day'] = '16'
        params['date_start_plus_month'] = '17'
        params['date_start_plus'] = '18'
        params['date_start_plus_year'] = '19'
        params['habaddr'] = '20'
        params['regaddr1'] = '21'
        params['regaddr2'] = '22'
        params['regaddr3'] = '23'
        params['inspect1'] = '24'
        params['inspect2'] = '25'
        params['is_AI'] = u"AI"
        params['is_AII'] = u"AII"
        params['is_AIII'] = u"AIII"
        params['is_AIV'] = u"AIV"
        params['is_B'] = u"B"
        params['is_C'] = u"C"
        params['is_D'] = u"D"
        params['is_E'] = u"E"
        params['is_F'] = u"F"
        params['#komment#'] = 'd'

        report.build(dst, params)

        self.assertEqual(os.path.exists(dst), True)


class TestUtils(unittest.TestCase):
    def test_date_to_float(self):
        """
        Тест преобразования даты в число
        """
        date_float = date_to_float(datetime(1899, 12, 30))
        self.assertEqual(date_float, 0)

        date_float = date_to_float(datetime(1899, 12, 31))
        self.assertEqual(date_float, 1)

        date_float = date_to_float(datetime(1899, 12, 29))
        self.assertEqual(date_float, 1)

        date_float = date_to_float(datetime(1899, 12, 29, 6))
        self.assertEqual(date_float, 1.25)

        date_float = date_to_float(datetime(1900, 1, 1))
        self.assertEqual(date_float, 2)

        date_float = date_to_float(datetime(1900, 1, 1, 6))
        self.assertEqual(date_float, 2.25)

if __name__ == '__main__':
    unittest.main()
