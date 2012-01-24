# coding: utf-8

from datetime import datetime

__author__ = 'prefer'

class TemplateTags(object):
    """
    Шаблонные теги

    Используется для того, чтобы задавать константы в шаблоне. Например, текущая дата
    будет задаваться через константу DATE_TIME, произвольные константы можно передовать
    через конструктор

    """
    def __init__(self, **kw):
        """
        Формируется словарь шаблонных тегов из переданных параметров
        """
        self.tags = {'DATE_NOW': datetime.now()}

        self.tags.update(kw)

    def get(self, key):
        """

        """
        return self.tags[key]