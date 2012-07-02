#coding: utf-8
from abc import ABCMeta, abstractmethod
from simple_report.utils import ZipProxy

__author__ = 'prefer'



class Document(object):
    """
    Базовый класс для всех документов
    """

    __metaclass__ = ABCMeta

    @abstractmethod
    def build(self):
        """
        """


class DocumentOpenXML(Document):
    u"""
    Базовый класс для работы со структурой open xml
    """

    __metaclass__ = ABCMeta

    def __init__(self, src_file, tags):
        self.extract_folder = ZipProxy.extract(src_file)

        self._tags = tags # Ссылка на тэги

    def pack(self, dst_file):
        """
        """
        self.build()
        ZipProxy.pack(dst_file, self.extract_folder)