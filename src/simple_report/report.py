#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

import abc
import os

from simple_report.interface import ISpreadsheetReport, IDocumentReport
from simple_report.converter.abstract import FileConverter
from simple_report.xlsx.document import DocumentXLSX
from simple_report.utils import FileProxy

class ReportGeneratorException(Exception):
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
            assert isinstance(converter, FileConverter)
            self.converter = converter

    def convert(self, src_file, to_format):
        """
        """
        if self.converter is not None:
            self.converter.set_src_file(src_file)
            return FileProxy(self.converter.build(to_format))
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

    @property
    def sections(self):
        return self.get_sections()

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

    @property
    def workbook(self):
        return self._wrapper.workbook

    @property
    def sheets(self):
        return self._wrapper.sheets


    def build(self, dst_file_path, file_type=XLSX):
        u"""
        Генерирует выходной файл в нужном формате

        @param dst_file_path: По этому пути будет находится результирующий файл
        @param file_type: Тип результирующего файла

        """
        if self.converter is None and file_type != self.XLSX:
            raise ReportGeneratorException('Converter is not defined')

        file_name, file_extension = os.path.splitext(dst_file_path)


        xlsx_path = os.path.extsep.join((file_name, self.XLSX))
        xlsx_file = FileProxy(xlsx_path, new_file=True)

        # Всегда вернет файл с расширением open office (xlsx, docx, etc.)

        self._wrapper.pack(xlsx_file)

        if file_type == self.XLSX:
            return xlsx_path
        else:
            return self.convert(xlsx_file, file_type)
