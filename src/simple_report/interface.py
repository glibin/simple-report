#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

from abc import ABCMeta, abstractmethod


class IReport(object):
    def show(self, *args, **kwargs):
        """
        Deprecated: use build
        """
        self.build(*args, **kwargs)

    @abstractmethod
    def build(self, *args, **kwargs):
        pass


class IDocumentReport(IReport):
    __metaclass__ = ABCMeta

    @abstractmethod
    def build(self, dst_file_path, params, file_type):
        u"""
        Генерирует выходной файл в нужном формате
        """

    @abstractmethod
    def get_all_parameters(self):
        u"""
        Возвращает все параметры документа
        """


class ISpreadsheetReport(IReport):
    __metaclass__ = ABCMeta

    @abstractmethod
    def get_sections(self):
        u"""
        Возвращает все секции
        """

    @abstractmethod
    def get_section(self, section_name):
        u"""
        Возвращает секцию по имени
        """

    @abstractmethod
    def build(self, dst_file_path, file_type):
        u"""
        Генерирует выходной файл в нужном формате
        """


class ISpreadsheetSection(object):
    __metaclass__ = ABCMeta


    # Тип разворота секции
    VERTICAL = 0
    HORIZONTAL = 1
    RIGHT_UP = 2
    LEFT_DOWN = 3
    RIGHT = 4
    HIERARCHICAL = 5

    @abstractmethod
    def flush(self, params, oriented=LEFT_DOWN):
        pass


    @abstractmethod
    def get_all_parameters(self):
        u"""
        Возвращает все параметры секции
        """