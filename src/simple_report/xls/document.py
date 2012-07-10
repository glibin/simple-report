#coding: utf-8

import xlrd

from simple_report.core.document_wrap import BaseDocument, SpreadsheetDocument
from simple_report.xls.workbook import Workbook
from simple_report.xls.output_options import XSL_OUTPUT_SETTINGS

class DocumentXLS(BaseDocument, SpreadsheetDocument):
    """
    """

    def __init__(self, ffile, tags=None, **kwargs):
        self.file = ffile
        self._workbook = Workbook(ffile, **kwargs)

    @property
    def workbook(self):
        return self._workbook

    def build(self, dst):
        self._workbook.build(dst)

    def __setattr__(self, key, value):

        if key in XSL_OUTPUT_SETTINGS:
            setattr(self._workbook, key, value)
        else:
            super(DocumentXLS, self).__setattr__(key, value)