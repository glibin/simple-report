#coding: utf-8
from abc import ABCMeta, abstractmethod, abstractproperty
from simple_report.utils import ZipProxy

__author__ = 'prefer'



class BaseDocument(object):
    """
    Базовый класс для всех документов
    """

    __metaclass__ = ABCMeta

    @abstractmethod
    def build(self):
        """
         Сборка документа
        :result:
        """


class SpreadsheetDocument(object):

    __metaclass__ = ABCMeta

    @abstractproperty
    def workbook(self):
        """
        Рабочая книга
        """

    def get_sections(self):
        """
        Возвращает все секции в шаблоне
        """
        return self.workbook.get_sections()

    def get_section(self, name):
        """
        Возвращает секцию по названию шаблона
        """
        return self.workbook.get_section(name)

    @property
    def sheets(self):
        """
        Листы отчета
        """
        return self.workbook.sheets


class DocumentOpenXML(BaseDocument):
    u"""
    Базовый класс для работы со структурой open xml
    """

    __metaclass__ = ABCMeta

    def __init__(self, src_file, tags):
        self.extract_folder = ZipProxy.extract(src_file)

        self._tags = tags # Ссылка на тэги

    def build(self, dst_file):
        """
        Сборка отчета
        """
        ZipProxy.pack(dst_file, self.extract_folder)
