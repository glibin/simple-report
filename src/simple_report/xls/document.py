#coding: utf-8

import xlrd

from simple_report.core.document_wrap import Document
from simple_report.xls.spreadsheet_ml import Workbook

class DocumentXLS(Document):
    """
    Делегат черной работы для XLSSpreadsheetReport
    """

    def __init__(self, ffile, tags=None):
        self.file = ffile
        self._workbook = Workbook(ffile)

    def get_section(self, name):
        return self._workbook.get_section(name)

    def get_sections(self):
        return self._workbook.get_sections()

    @property
    def workbook(self):
        return self._workbook

    @property
    def sheets(self):
        return self._workbook._sheet_list()

    def build(self):
        """
        """

    def show(self, dst, file_type):
        self._workbook.show(dst, file_type)