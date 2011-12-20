#coding: utf-8

from simple_report.converter.abstract import FileConverter, FileConverterException
from simple_report.converter.open_office.wrapper import OOWrapper

__author__ = 'prefer'

class OpenOfficeConverter(FileConverter):
    """

    """

    def convert(self, to_format):
        u"""
        Метод должен исходя из исходного типа документа и требуемого типа
        найти метод у себя и вызвать его. Если этого метода нет - должно
        генериться исключение.
        """
        if to_format in (FileConverter.XLSX, FileConverter.DOCX):
            raise FileConverterException('Format "%s" not supported' % to_format)
        return OOWrapper().convert(self.file.get_path(), to_format)

