#coding: utf-8
from simple_report.core.document_wrap import DocumentOpenXML
from simple_report.xlsx.spreadsheet_ml import CommonPropertiesXLSX

__author__ = 'prefer'

class DocumentXLSX(DocumentOpenXML):
    u"""
    """

    def __init__(self, *args, **kwargs):
        super(DocumentXLSX, self).__init__(*args, **kwargs)
        self.common_properties = CommonPropertiesXLSX.create(self.extract_folder, self._tags)

    def get_sections(self):
        u"""
        Возвращает все секции в шаблоне
        """
        return self.workbook.get_sections()

    def get_section(self, name):
        u"""
        Возвращает секцию по названию шаблона
        """
        return self.workbook.get_section(name)

    @property
    def workbook(self):
        return self.common_properties.main

    @property
    def sheets(self):
        return self.workbook.sheets

    def build(self):
        """
        """
        self.workbook.build()

