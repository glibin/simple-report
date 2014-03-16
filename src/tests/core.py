#coding: utf-8

from datetime import datetime
import sys
import os
import subprocess
sys.path.append('/home/vahotin/dev/simple_report/src')
sys.path.append(os.path.join(os.path.dirname(__file__), '../'))
from simple_report.converter.abstract import FileConverter
from simple_report.core.tags import TemplateTags
from simple_report.interface import ISpreadsheetSection
from simple_report.xlsx.section import Section, MergeXLSX
from simple_report.xlsx.spreadsheet_ml import (SectionException,
                                               SectionNotFoundException)
from simple_report.xlsx.formula import Formula
from simple_report.utils import ColumnHelper, date_to_float, FormulaWriteExcel
from simple_report.xls.section import XLSImage

sys.path.append('.')

from test_oo_wrapper import TestOO
from test_utils import skip_python26
from test_pko import TestPKO
from oborot import OperationsJournalReportFactory
from test_pagebreaks import TestPagebreaks

sys.path.append('../')

import os
import unittest

from simple_report.report import (SpreadsheetReport, ReportGeneratorException,
                                  DocumentReport)
from simple_report.xls.document import DocumentXLS
from simple_report.xls.section import MergeXLS


LOREM_IPSUM = (
    'Lorem ipsum dolor sit amet, consectetur adipisicing elit, '
    'sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '
    'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris '
    'nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in '
    'reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla '
    'pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa '
    'qui officia deserunt mollit anim id est laborum.'
)


class TestXLSX(object):
    """

    """

    # Разные директории с разными файлами под linux и под windows
    SUBDIR = None

    def setUp(self):
        assert self.SUBDIR
        self.src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    'test_data', self.SUBDIR, 'xlsx', )
        self.dst_dir = self.src_dir

        self.test_files = dict([(path, os.path.join(self.src_dir, path))
                                for path in os.listdir(self.src_dir)
                                if path.startswith('test')])

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


class TestLinuxXLSX(TestXLSX, TestOO, TestPKO, TestPagebreaks,
                    unittest.TestCase):
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

    def test_formula_generation(self):
        src = self.test_files['test-formula_generation.xlsx']
        dst = os.path.join(self.dst_dir, 'res-formula_generation.xlsx')
        if os.path.exists(dst):
            os.remove(dst)

        report = SpreadsheetReport(src)

        all_change = report.get_section('all_change')
        all_not_change = report.get_section('all_not_change')
        row_not_change = report.get_section('row_not_change')
        column_not_change = report.get_section('column_not_change')
        other_section = report.get_section('other_section')
        row_insert_formula_section = report.get_section('row_insert_formula')
        check_insert_formula_section = report.get_section(
            'check_insert_formula')

        all_change_formula = '(A1+B1)*3'
        all_not_change_formula = '($A$1+B1)*3'
        row_not_change_formula = '(A$1+B1)*3'
        column_not_change_formula = '($A1+B1)*3'

        all_change.flush({'p1': 1, 'p2': 1},
                         oriented=ISpreadsheetSection.VERTICAL)
        all_change.flush({'p1': 2, 'p2': 2},
                         oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(Formula.get_instance(all_change_formula).formula,
                         '(A2+B2)*3')
        other_section.flush({'p1': 1}, oriented=ISpreadsheetSection.VERTICAL)
        all_change.flush({'p1': 3, 'p2': 3},
                         oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(Formula.get_instance(all_change_formula).formula,
                         '(A4+B4)*3')

        other_section.flush({'p1': 1}, oriented=ISpreadsheetSection.VERTICAL)
        all_not_change.flush({'p1': 1, 'p2': 1},
                             oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(Formula.get_instance(all_not_change_formula).formula,
                         '($A$1+B1)*3')
        all_not_change.flush({'p1': 2, 'p2': 2},
                             oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(Formula.get_instance(all_not_change_formula).formula,
                         '($A$1+B2)*3')
        other_section.flush({'p1': 1},
                            oriented=ISpreadsheetSection.VERTICAL)
        all_not_change.flush({'p1': 3, 'p2': 3},
                             oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(Formula.get_instance(all_not_change_formula).formula,
                         '($A$1+B4)*3')

        other_section.flush({'p1': 1},
                            oriented=ISpreadsheetSection.VERTICAL)
        row_not_change.flush({'p1': 1, 'p2': 1},
                             oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(Formula.get_instance(row_not_change_formula).formula,
                         '(A$1+B1)*3')
        row_not_change.flush({'p1': 2, 'p2': 2},
                             oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(Formula.get_instance(row_not_change_formula).formula,
                         '(A$1+B2)*3')
        other_section.flush({'p1': 1}, oriented=ISpreadsheetSection.VERTICAL)
        row_not_change.flush({'p1': 3, 'p2': 3},
                             oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(Formula.get_instance(row_not_change_formula).formula,
                         '(A$1+B4)*3')

        other_section.flush({'p1': 1}, oriented=ISpreadsheetSection.VERTICAL)
        column_not_change.flush({'p1': 1, 'p2': 1},
                                oriented=ISpreadsheetSection.VERTICAL)
        self.assertEqual(
            Formula.get_instance(
                column_not_change_formula
            ).formula, '($A1+B1)*3'
        )
        other_section.flush({'p1': 1}, oriented=ISpreadsheetSection.HORIZONTAL)
        column_not_change.flush({'p1': 2, 'p2': 2},
                                oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(
            Formula.get_instance(
                column_not_change_formula
            ).formula, '($A1+F1)*3'
        )
        other_section.flush({'p1': 1}, oriented=ISpreadsheetSection.HORIZONTAL)
        column_not_change.flush({'p1': 3, 'p2': 3},
                                oriented=ISpreadsheetSection.HORIZONTAL)
        self.assertEqual(
            Formula.get_instance(
                column_not_change_formula
            ).formula, '($A1+J1)*3'
        )

        insert_formulas(row_insert_formula_section,
                        check_insert_formula_section)
        # Проверяем, что вписанные формулы попали в дерево и правильно
        # записались
        found_B21 = found_C21 = False
        for row in report.sheets[0].sheet_data.write_data.getchildren():
            if row.tag == 'row':
                for c_ in row.getchildren():
                    if c_.tag == 'c':
                        if c_.attrib.get('r') == "B21":
                            found_B21 = True
                            func = c_.find('f')
                            assert func is not None
                            assert func.text == 'AVERAGE(B17:B20)'
                        elif c_.attrib.get('r') == "C21":
                            found_C21 = True
                            func = c_.find('f')
                            assert func is not None
                            assert func.text == 'SUM((A17,A18,A19,A20))'
        assert found_B21 and found_C21

        return report.build(dst)

    def test_with_merge(self):
        """
        Конструкция with merge для обьединения ячеек вывода
        """

        src = self.test_files['test_merge.xlsx']
        dst = os.path.join(self.dst_dir, 'res-merge.xlsx')
        if os.path.exists(dst):
            os.remove(dst)

        report = SpreadsheetReport(src)
        s1 = report.get_section('s1')
        s2 = report.get_section('s2')
        s3 = report.get_section('s3')
        s4 = report.get_section('s4')
        s5 = report.get_section('s5')

        s5.flush({'p5': 1}, oriented=ISpreadsheetSection.VERTICAL)
        s5.flush({'p5': 2}, oriented=ISpreadsheetSection.VERTICAL)
        s5.flush({'p5': 3}, oriented=ISpreadsheetSection.HORIZONTAL)

        m1 = MergeXLSX(s1, s2, {'p1': 1},
                       oriented=ISpreadsheetSection.HORIZONTAL)
        with m1:
            with MergeXLSX(s2, s3, {'p21': 1, 'p22': 21},
                           oriented=ISpreadsheetSection.HORIZONTAL):
                m3 = MergeXLSX(s3, s4, {'p3': 1},
                               oriented=ISpreadsheetSection.HORIZONTAL)
                with m3:
                    s4.flush({'p4': 1},
                             oriented=ISpreadsheetSection.RIGHT)
                    for i in range(2, 4):
                        s4.flush({'p4': i},
                                 oriented=ISpreadsheetSection.VERTICAL)

                m3_exp = (m3._begin_merge_col == 'J'
                          and m3._end_merge_col == 'J'
                          and m3.begin_row_merge == 4
                          and m3.end_row_merge == 6)
                self.assertEqual(m3_exp, True)

                with MergeXLSX(s3, s4, {'p3': 2},
                               oriented=ISpreadsheetSection.HIERARCHICAL):
                    s4.flush({'p4': 1}, oriented=ISpreadsheetSection.RIGHT)
                    s4.flush({'p4': 2}, oriented=ISpreadsheetSection.VERTICAL)

            with MergeXLSX(s2, s3, {'p21': 2, 'p22': 21},
                           oriented=ISpreadsheetSection.HIERARCHICAL):
                with MergeXLSX(s3, s4, {'p3': 1},
                               oriented=ISpreadsheetSection.HORIZONTAL):
                    s4.flush({'p4': 1}, oriented=ISpreadsheetSection.RIGHT)
                    s4.flush({'p4': 2}, oriented=ISpreadsheetSection.VERTICAL)

        m1_exp = (m1._begin_merge_col == 'G' and m1._end_merge_col == 'G'
                  and m1.begin_row_merge == 4 and m1.end_row_merge == 10)
        self.assertEqual(m1_exp, True)

        with MergeXLSX(s1, s2, {'p1': 2},
                       oriented=ISpreadsheetSection.HIERARCHICAL):
            with MergeXLSX(s2, s3, {'p21': 1, 'p22': 21},
                           oriented=ISpreadsheetSection.HORIZONTAL):
                with MergeXLSX(s3, s4, {'p3': 1},
                               oriented=ISpreadsheetSection.HORIZONTAL):
                    s4.flush({'p4': 1},
                             oriented=ISpreadsheetSection.HORIZONTAL)

        return report.build(dst)

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

    def test_simple_example(self):
        """

        """

        file_name = 'test_simple.xlsx'
        template_name = self.test_files[file_name]
        report = SpreadsheetReport(template_name)

        header = report.get_section('header')
        header.flush({
            'author': u'BARS Group',
            'today': datetime.now().strftime('%d.%m.%Y')
        })

        res_file_name = 'res_simple.xlsx'
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
                                            'fio': u'Иванов %d' % i,
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
        self.src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    'test_data', self.SUBDIR, 'xlsx', )
        self.dst_dir = self.src_dir

        self.test_files = dict([(path, os.path.join(self.src_dir, path))
                                for path in os.listdir(self.src_dir)
                                if path.startswith('test')])

    def test_left_down(self, report=None):
        if report is None:
            return
        for i in range(2):
            section1 = report.get_section('Section1')
            section1.flush({'section1': i}, oriented=Section.LEFT_DOWN)
            self.assertEqual(section1.sheet_data.cursor.row, ('A', 2 * i + 3))
            self.assertEqual(section1.sheet_data.cursor.column,
                             ('C', 2 * i + 1))

    def test_left_down2(self, report=None):
        if report is None:
            return
        for i in range(2):
            section3 = report.get_section('Section3')
            section3.flush({'section3': 100}, oriented=Section.LEFT_DOWN)
            self.assertEqual(section3.sheet_data.cursor.row, ('A', 2 * i + 11))
            self.assertEqual(section3.sheet_data.cursor.column,
                             ('C', 2 * i + 9))

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
            self.assertEqual(section2.sheet_data.cursor.row, ('C', 2 * i + 5))
            self.assertEqual(section2.sheet_data.cursor.column,
                             ('E', 2 * i + 3))

    def test_horizontal(self, report=None):
        if report is None:
            return
        for i in range(3):
            section3 = report.get_section('Section3')
            section3.flush({'section3': i}, oriented=Section.HORIZONTAL)
            self.assertEqual(section3.sheet_data.cursor.row, ('C', 9))
            self.assertEqual(section3.sheet_data.cursor.column,
                             (ColumnHelper.add('G', 2 * i), 7))

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




class TestWindowsXLSX(TestXLSX, unittest.TestCase):
    SUBDIR = 'win'


class TestLinuxDOCX(unittest.TestCase):
    """

    """

    def setUp(self):
        self.src_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            'test_data',
            'linux',
            'docx'
        )
        self.dst_dir = self.src_dir

        self.test_files = dict([(path, os.path.join(self.src_dir, path))
                                for path in os.listdir(self.src_dir)
                                if path.startswith('test')])

    def test_simple_docx(self):
        """
        """

        template_name = 'test-sluzh.docx'
        path = self.test_files[template_name]
        doc = DocumentReport(path)

        res_file_name = 'res-' + template_name
        dst = os.path.join(self.dst_dir, res_file_name)

        doc.build(dst, {'Employee_name': u'Иванова И.И.',
                        'region_name': u'Казань'})
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

        self.assertFalse(tags[0] != '#brandgroupname#'
                         and tags[0] != '#category#')
        self.assertFalse(tags[1] != '#brandgroupname#'
                         and tags[1] != '#category#')

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

    def test_tables_flush(self):
        template_name = 'test_table.docx'
        path = self.test_files[template_name]

        res_file_name = 'res-table_flush.docx'
        dst = os.path.join(self.dst_dir, res_file_name)

        report = DocumentReport(path)
        # report.set_docx_table_sections()
        s1 = report.get_section('section1')
        s2 = report.get_section('section2')
        s2.flush({'test': u'Lorem ipsum'})
        s1.flush({
            'test_table_row1col1': u'Hello',
            'test_table_row1col2': u'simple_report',
            'test_table_row1col3': u'user',
            'test_table_row1col4': LOREM_IPSUM,
        })
        params = {}
        report.build(dst, params)

    def test_table_section_double(self):
        template_name = 'test_table_double_section.docx'
        path = self.test_files[template_name]

        report = DocumentReport(path)
        try:
            report.get_section('section1')
        except SectionException:
            pass
        else:
            raise Exception('Docx tables sections doubling test failed')

    def test_divisible_keys(self):
        template_name = 'test_divisibles.docx'
        path = self.test_files[template_name]
        report = DocumentReport(path)
        res_file_name = 'res-divisibles.docx'
        dst = os.path.join(self.dst_dir, res_file_name)
        params = {
            "tasks": "",
            "kind_tostring": u"документарная и выездная",
            "normative_list": "",
            "finish_date": "13.12.2012",
            "expert_list": "",
            "docs": "",
            "num": "1",
            "purpose": "",
            "address": u"420101, Респ Татарстан (Татарстан), г Казань, ул Карбышева, д. 37, кв. 44",
            "events": "",
            "subject3": "x",
            "articles": "",
            "inspectors_list": "",
            "supervisionobj_name": u"Малыши и малышки",
            "oyear": 2013,
            "type_tostring": u"внеплановая",
            "start_date": "14.02.2013",
            "subject1": "x",
            "subject2": "x",
            "oday": 21,
            "subject4": "x",
            "subject5": "x",
            "subject6": "x",
            "subject7": "x",
            "authority_parent": "",
            "omonth": 3
        }
        report.build(dst, params)

    def test_flush_order(self):
        template_name = 'test_flush_order.docx'
        path = self.test_files[template_name]
        report = DocumentReport(path)
        res_file_name = 'res-flush_order.docx'
        dst = os.path.join(self.dst_dir, res_file_name)
        params = {
            "example": "output_one",
            "example_two": "ouput_two",
            "example_two_three": "output_two_three",
            "exampl": "no_output"
        }
        report.build(dst, params)

    def test_tabs(self):
        template_name = 'test_tabs.docx'
        path = self.test_files[template_name]
        report = DocumentReport(path)
        res_file_name = 'res-tabs.docx'
        dst = os.path.join(self.dst_dir, res_file_name)
        params = {
            "tfoms_to": "TFOMS",
            "tfoms_to_address": "TFOMS_ADDRESS",
            "tfoms_to_director_fio": "TFOMS_TO_DIR_FIO"
        }
        report.build(dst, params)


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


def insert_formulas(row_insert_formula_section, check_insert_formula_section):
    for j in range(4):
        row_insert_formula_section.flush(
            {'p1': 5 + j, 'p2': 6 - j, 'p3': 4 * j},
            used_formulas={'p1': ['p1', 't1'], 'p2': ['p2']}
        )
    check_insert_formula_section.flush(
        {
            'f2': FormulaWriteExcel('p2', 'AVERAGE', True),
            'f3': FormulaWriteExcel('p1', 'SUM', False)
        },
        oriented=ISpreadsheetSection.VERTICAL
    )

if __name__ == '__main__':
    unittest.main()

