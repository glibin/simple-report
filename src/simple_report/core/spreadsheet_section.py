#coding: utf-8

import abc

class SpreadsheetSection(object):

    __metaclass__ = abc.ABCMeta

    def __init__(self, sheet, name, begin=None, end=None):
        """
        Абстракная секция для таблиц
        """

        assert begin or end

        self.name = name
        self.begin = begin
        self.end = end