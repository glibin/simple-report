#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

import abc

from interface import ISpreadsheetReport, IDocumentReport
from simple_report.converter.abstract import FileConverter
from simple_report.xlsx.document import DocumentXLSX
from utils import FileProxy

class ReportException(Exception):
    """
    """

class Report(object):
    u"""
    Абстрактный класс отчета
    """

    __metaclass__ = abc.ABCMeta

    def __init__(self, src_file, converter=None):
        """
        """
        self.file = FileProxy(src_file)

        self.converter = None
        if converter is not None:

            assert issubclass(converter, FileConverter)
            self.converter = converter


    def convert(self, src_file, to_format):
        """
        """
        if self.converter is not None:
            return FileProxy(self.converter(src_file).build(to_format))
        else:
            return src_file


class DocumentReport(Report, IDocumentReport):
    DOCX = FileConverter.DOCX

    def build(self, dst_file_path, params, file_type=DOCX):
        u"""
        Генерирует выходной файл в нужном формате
        """


class SpreadsheetReport(Report, ISpreadsheetReport):
    XLSX = FileConverter.XLSX

    def __init__(self, *args, **kwargs):
        super(SpreadsheetReport, self).__init__(*args, **kwargs)

        xlsx_file = self.convert(self.file, self.XLSX)
        self._wrapper = DocumentXLSX(xlsx_file)

    def get_sections(self):
        u"""
        Возвращает все секции
        """

        return self._wrapper.get_sections()

    def get_section(self, section_name):
        u"""
        Возвращает секцию по имени
        """
        return self._wrapper.get_section(section_name)


    def build(self, dst_file_path, file_type=XLSX):
        u"""
        Генерирует выходной файл в нужном формате
        """
        if self.converter is None and file_type != self.XLSX:
            raise ReportException('Converter is not defined')

        dst_file = FileProxy(dst_file_path, new_file=True)
        self._wrapper.pack(dst_file)
        return self.convert(dst_file, file_type)
