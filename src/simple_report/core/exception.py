#coding: utf-8

import abc

class AbstractSectionException(Exception):
    """
    Абстрактный класс для исключений, которые возникают при работе с секциями
    """

    __metaclass__ = abc.ABCMeta

class SectionException(AbstractSectionException):
    """
    """


class SectionNotFoundException(AbstractSectionException):
    """
    Исключение - секция не найдена
    """


class SheetException(Exception):
    """
    Абстрактный класс для исключений, которые возникают при работе с листами таблицы
    """

    __metaclass__ = abc.ABCMeta


class SheetNotFoundException(SheetException):
    """
    Исключение лист не найден
    """


class SheetDataException(SheetException):
    """
    Ошибка данных.
    """

class XLSReportWriteException(Exception):
    """
    Ошибка вывода в отчетах XLS
    """
