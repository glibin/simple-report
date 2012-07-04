#coding: utf-8

import xlrd

from simple_report.core.document_wrap import BaseDocument, SpreadsheetDocument
from simple_report.xls.workbook import Workbook

class DocumentXLS(BaseDocument, SpreadsheetDocument):
    """
    """

    def __init__(self, ffile, tags=None):
        self.file = ffile
        self._workbook = Workbook(ffile)

    @property
    def workbook(self):
        return self._workbook

    def build(self, dst):
        self._workbook.build(dst)