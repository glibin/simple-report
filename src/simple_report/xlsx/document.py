#coding: utf-8
from simple_report.core.document_wrap import DocumentOpenXML, SpreadsheetDocument
from simple_report.xlsx.spreadsheet_ml import CommonPropertiesXLSX

__author__ = 'prefer'

class DocumentXLSX(DocumentOpenXML, SpreadsheetDocument):
    u"""
    """

    def __init__(self, *args, **kwargs):
        super(DocumentXLSX, self).__init__(*args, **kwargs)
        self.common_properties = CommonPropertiesXLSX.create(self.extract_folder, self._tags)

    @property
    def workbook(self):
        return self.common_properties.main

    def build(self, dst_file):
        """
        """
        self.workbook.build()
        super(DocumentXLSX, self).build(dst_file)
