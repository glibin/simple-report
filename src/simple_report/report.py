#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

import abc
import os
from simple_report.core.document_wrap import BaseDocument
from simple_report.core.exception import WrongDocumentType
from simple_report.core.tags import TemplateTags
from simple_report.docx.document import DocumentDOCX

from simple_report.interface import ISpreadsheetReport, IDocumentReport
from simple_report.converter.abstract import FileConverter
from simple_report.xlsx.document import DocumentXLSX
from simple_report.xls.document import DocumentXLS
from simple_report.xls.output_options import XSL_OUTPUT_SETTINGS
from simple_report.utils import FileProxy




class ReportGeneratorException(Exception):
    """
    """


class Report(object):
    u"""
    Абстрактный класс отчета
    """

    __metaclass__ = abc.ABCMeta

    def __init__(self, src_file, converter=None, tags=None, wrapper=None,
                 type=None, **kwargs):
        """
        """

        self.TYPE = type

        self.tags = tags or TemplateTags()
        assert isinstance(self.tags, TemplateTags)

        self.file = FileProxy(src_file)

        self.converter = None
        if converter is not None:
            assert isinstance(converter, FileConverter)
            self.converter = converter

        ffile = self.convert(self.file, self.TYPE)

        assert issubclass(wrapper, BaseDocument)
        self._wrapper = wrapper(ffile, self.tags, **kwargs)

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

        self._wrapper.build(xlsx_file)

        if file_type == self.TYPE:
            return xlsx_path
        else:
            return self.convert(xlsx_file, file_type)


class DocumentReport(Report, IDocumentReport):
    #

    def __init__(self, src_file, converter=None, tags=None, wrapper=DocumentDOCX, type=FileConverter.DOCX):

        assert issubclass(wrapper, DocumentDOCX)
        assert (type == FileConverter.DOCX)

        super(DocumentReport, self).__init__(src_file, converter, tags, wrapper, type)

    def build(self, dst_file_path, params, file_type=FileConverter.DOCX):
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

    def get_section(self, section_name):
        return self._wrapper.get_section(section_name)


class SpreadsheetReport(Report, ISpreadsheetReport):

    def __init__(self, src_file, converter=None, tags=None, wrapper=DocumentXLSX, type=FileConverter.XLSX, **kwargs):

        assert issubclass(wrapper, DocumentXLSX) or issubclass(wrapper, DocumentXLS)
        assert (type == FileConverter.XLSX) or (type == FileConverter.XLS)

        super(SpreadsheetReport, self).__init__(src_file, converter, tags, wrapper, type, **kwargs)

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
        if not hasattr(self._wrapper, 'get_section'):
            raise WrongDocumentType()
        return self._wrapper.get_section(section_name)

    @property
    def workbook(self):
        return self._wrapper.workbook

    @property
    def sheets(self):
        return self._wrapper.sheets

    def __setattr__(self, key, value):

        if key in XSL_OUTPUT_SETTINGS:
            setattr(self._wrapper, key, value)
        else:
            super(SpreadsheetReport, self).__setattr__(key, value)
