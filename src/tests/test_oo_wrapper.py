# coding: utf-8
import os
from simple_report.converter.abstract import FileConverterException, FileConverter
from simple_report.converter.open_office import OOWrapper, settings, OpenOfficeConverter
from simple_report.converter.open_office.wrapper import OOWrapperException
from simple_report.report import SpreadsheetReport
from simple_report.xlsx.section import Section

class TestOO(object):
    def test_connection(self):
        try:
            OOWrapper()
        except OOWrapperException as e:
            # self.fail('OOWrapper() raised OOWrapperException - %s ' % str(e))
            print u'Тесты для OpenOffice не были запущены, '\
                  u'т.к. нет запущенного сервера OpenOffice на порту %s ' % settings.DEFAULT_OPENOFFICE_PORT

            return False
        else:
            return True

    def test_oo_wrapper(self):
        """
        Тестирование OpenOffice конвертера
        """
        if not self.test_connection():
             return

        src = self.test_files['test-PF_PKO.xlsx']

        converter = OOWrapper()
        file_path = converter.convert(src, 'xls')
        self.assertEqual(os.path.exists(file_path), True)

        with  self.assertRaises(OOWrapperException):
            converter.convert(src, 'odt') # Для Writera


    def test_work_document(self):
        if not self.test_connection():
            return

        with self.assertRaises(FileConverterException):
            src = self.test_files['test-simple.xls']
            report = SpreadsheetReport(src, converter=OpenOfficeConverter)

        src = self.test_files['test-simple.xlsx']
        report = SpreadsheetReport(src, converter=OpenOfficeConverter)
        dst = os.path.join(self.dst_dir, 'res-simple.xlsx')

        if os.path.exists(dst):
            os.remove(dst)
        self.assertEqual(os.path.exists(dst), False)



        self.assertGreater(len(report._wrapper.sheets), 0)
        self.assertLessEqual(len(report._wrapper.sheets), 4)

        self.assertNotEqual(report._wrapper.workbook, None)
        self.assertNotEqual(report._wrapper.workbook.shared_strings, None)

        # Тестирование получения секции
        section_a1 = report.get_section('A1')
        self.assertIsInstance(section_a1, Section)

        with self.assertRaises(Exception):
            report.get_section('G1')

        section_a1.flush({'user': u'Иванов Иван',
                          'date_now': 1})
        for i in range(10):
            report.get_section('B1').flush({'nbr': i,
                                            'fio': u'Иванов %d' % i})

        report.get_section('C1').flush({'user': u'Иван'})
        report.build(dst, FileConverter.XLS)

        self.assertEqual(os.path.exists(dst), True)