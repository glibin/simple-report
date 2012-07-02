#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

import abc
import os
from simple_report.core.document_wrap import Document, DocumentOpenXML
from simple_report.core.tags import TemplateTags
from simple_report.docx.document import DocumentDOCX

from simple_report.interface import ISpreadsheetReport, IDocumentReport
from simple_report.converter.abstract import FileConverter
from simple_report.xlsx.document import DocumentXLSX
from simple_report.xls.document import DocumentXLS
from simple_report.utils import FileProxy

class ReportGeneratorException(Exception):
    """
    """


class Report(object):
    u"""
    Абстрактный класс отчета
    """

    __metaclass__ = abc.ABCMeta

    # Тип документа: XLSX, DOCX, etc.
    TYPE = None

    # Класс-делегат черной работы
    _wrapper = None

    def __init__(self, src_file, converter=None, tags=None):
        """
        """

        self.tags = tags or TemplateTags()
        assert isinstance(self.tags, TemplateTags)

        self.file = FileProxy(src_file)

        self.converter = None
        if converter is not None:
            assert isinstance(converter, FileConverter)
            self.converter = converter

        ffile = self.convert(self.file, self.TYPE)

        assert issubclass(self._wrapper, Document)
        self._wrapper = self._wrapper(ffile, self.tags)

    def convert(self, src_file, to_format):
        """
        """
        if self.converter is not None:
            self.converter.set_src_file(src_file)
            return FileProxy(self.converter.build(to_format))
        else:
            return src_file

    def build(self, dst_file_path, file_type=None):
        """

        """
        assert self.TYPE, 'Document Type is not defined'

        if file_type is None:
            file_type = self.TYPE

        if self.converter is None and file_type != self.TYPE:
            raise ReportGeneratorException('Converter is not defined')

        file_name, file_extension = os.path.splitext(dst_file_path)

        xlsx_path = os.path.extsep.join((file_name, self.TYPE))
        xlsx_file = FileProxy(xlsx_path, new_file=True)

        # Всегда вернет файл с расширением open office (xlsx, docx, etc.)

        self._wrapper.pack(xlsx_file)

        if file_type == self.TYPE:
            return xlsx_path
        else:
            return self.convert(xlsx_file, file_type)


class DocumentReport(Report, IDocumentReport):
    #
    TYPE = FileConverter.DOCX

    _wrapper = DocumentDOCX


    def build(self, dst_file_path, params, file_type=TYPE):
        u"""
        Генерирует выходной файл в нужном формате
        """
        self._wrapper.set_params(params)
        return super(DocumentReport, self).build(dst_file_path, file_type)

    def get_all_parameters(self):
        """
        Возвращаем параметры отчета.
        """
        return self._wrapper.get_all_parameters()


class SpreadsheetReport(Report, ISpreadsheetReport):
    TYPE = FileConverter.XLSX

    _wrapper = DocumentXLSX

    def __init__(self, src_file, converter=None, tags=None, wrapper=DocumentXLSX, type=FileConverter.XLSX):

        self.TYPE = type
        self._wrapper = wrapper

        super(SpreadsheetReport, self).__init__(src_file, converter, tags)

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

    def build(self, dst_file_path, file_type=None):

        if isinstance(self._wrapper, DocumentXLSX):
            super(SpreadsheetReport, self).build(dst_file_path, file_type=None)
        else:
            self._wrapper.show(dst=dst_file_path, file_type='xls')


class XLSSpreadsheetReport(Report, ISpreadsheetReport):

    TYPE = FileConverter.XLS

    _wrapper = DocumentXLS

    def get_section(self, section_name):
        return self._wrapper.get_section(section_name)

    def get_sections(self):
        return self._wrapper.get_sections()

    @property
    def workbook(self):
        return self._wrapper.workbook

    @property
    def sheets(self):
        return self._wrapper.sheets

    def build(self, dst, type='xls'):
        self._wrapper.show(dst=dst, file_type=type)