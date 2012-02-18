#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''
from simple_report.core.document_wrap import DocumentOpenXML
from simple_report.docx.wordprocessing_ml import CommonPropertiesDOCX


class DocumentDOCX(DocumentOpenXML):
    """
    """

    def __init__(self, *args, **kwargs):
        super(DocumentDOCX, self).__init__(*args, **kwargs)
        self.common_properties = CommonPropertiesDOCX.create(self.extract_folder, self._tags)

    @property
    def word(self):
        return self.common_properties.main

    def build(self):
        """
        """
        self.word.build()

    def set_params(self, *args, **kwargs):
        """

        """
        self.word.set_params(*args, **kwargs)