#coding: utf-8
from abc import ABCMeta, abstractmethod
from simple_report.utils import ZipProxy

__author__ = 'prefer'


class DocumentOpenXML(object):
    u"""
    Базовый класс для работы со структурой open xml
    """

    __metaclass__ = ABCMeta

    def __init__(self, src_file):
        self.extract_folder = ZipProxy.extract(src_file)

    def pack(self, dst_file):
        """
        """
        self.build()
        ZipProxy.pack(dst_file, self.extract_folder)

    @abstractmethod
    def build(self):
        """

        """