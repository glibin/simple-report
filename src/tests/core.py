#coding: utf-8
from datetime import datetime

import sys;
from simple_report.converter.abstract import FileConverter
from simple_report.core.tags import TemplateTags
from simple_report.interface import ISpreadsheetSection
from simple_report.xlsx.section import Section
from simple_report.xlsx.spreadsheet_ml import SectionException, SectionNotFoundException

from test_oo_wrapper import TestOO
from test_utils import skip_python26
from test_pko import TestPKO
from oborot import OperationsJournalReportFactory

sys.path.append('../')

import os
import unittest

from simple_report.report import SpreadsheetReport, ReportGeneratorException, DocumentReport
from simple_report.utils import ColumnHelper


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


class TestLinuxXLSX(TestXLSX, TestOO, TestPKO, unittest.TestCase):
    SUBDIR = 'linux'

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

if __name__ == '__main__':
    unittest.main()
