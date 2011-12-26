#coding: utf-8

import sys;
from simple_report.converter.abstract import FileConverter
from simple_report.core.shared_table import SharedStringsTable
from simple_report.xlsx.section import Section
from simple_report.xlsx.spreadsheet_ml import SectionException, SectionNotFoundException, WorkbookSheet

from test_oo_wrapper import TestOO
from test_utils import skip_python26
from test_pko import TestPKO

sys.path.append('../')

import os
import unittest

from simple_report.report import SpreadsheetReport, ReportException
from simple_report.utils import ColumnHelper




class SetupData(object):
    """

    """

    # Разные директории с разными файлами под linux и под windows
    SUBDIR = None

    def setUp(self):
        assert self.SUBDIR
        self.src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data', self.SUBDIR)
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

        report = SpreadsheetReport(src)

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
        s_gor.flush({'col': u'Данные'}, oriented=s_gor.GORIZONTAL)

        for i in range(10):
            report.get_section('B1').flush({'nbr': i,
                                            'fio': u'Иванов %d' % i,
                                            'sector': u'Какой-то сектор'})

            s_gor_str = report.get_section('GorStr')
            s_gor_str.flush({'g': i+i}, oriented=s_gor.GORIZONTAL)
            s_gor_str.flush({'g': i*i}, oriented=s_gor.GORIZONTAL)


        report.get_section('C1').flush({'user': u'Иван'})

        with self.assertRaises(ReportException):
            report.build(dst, FileConverter.XLS)

        report.build(dst)

        self.assertEqual(os.path.exists(dst), True)


    def test_workbook_with_2_6_python(self):

        src = self.test_files['test-simple.xlsx']
        dst = os.path.join(self.dst_dir, 'res-simple.xlsx')
        if os.path.exists(dst):
            os.remove(dst)
        self.assertEqual(os.path.exists(dst), False)

        report = SpreadsheetReport(src)

        section_a1 = report.get_section('A1')
        section_a1.flush({'user': u'Иванов Иван',
                          'date_now': 1})

        s_gor = report.get_section('GOR')
        s_gor.flush({'col': u'Данные'}, oriented=s_gor.GORIZONTAL)

        for i in range(10):
            report.get_section('B1').flush({'nbr': i,
                                            'fio': u'Иванов %d' % i,
                                            'sector': u'Какой-то сектор'})

            s_gor_str = report.get_section('GorStr')
            s_gor_str.flush({'g': i+i}, oriented=s_gor.GORIZONTAL)
            s_gor_str.flush({'g': i*i}, oriented=s_gor.GORIZONTAL)


        report.get_section('C1').flush({'user': u'Иван'})
        report.build(dst)
        self.assertEqual(os.path.exists(dst), True)

class TestLinux(SetupData, TestOO, TestPKO,  unittest.TestCase):
    SUBDIR = 'linux'

    def test_fake_section(self):
        src = self.test_files['test-simple-fake-section.xlsx']
        with self.assertRaises(SectionException):
            report = SpreadsheetReport(src)


    def test_merge_cells(self):
        src = self.test_files['test-merge-cells.xlsx']
        dst = os.path.join(self.dst_dir, 'res-merge-cells.xlsx')


        report = SpreadsheetReport(src)

        report.get_section('head').flush({'kassa_za':u'Ноябрь'})

        for i in range(10):
            report.get_section('table_dyn').flush({'doc_num': i})


        report.get_section('foot').flush({'glavbuh':u'Иван'})

        report.build(dst)

class TestWindows(SetupData, unittest.TestCase):
    SUBDIR = 'win'


if __name__ == '__main__':
    unittest.main()
