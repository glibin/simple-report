#coding: utf-8

import abc

class AbstractSectionException(Exception):
    """
    Абстрактный класс для исключений, которые возникают при работе с секциями
    """

    __metaclass__ = abc.ABCMeta

class SectionException(AbstractSectionException):
    """
    Исключение работы с секциями
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
    Исключение "Лист не найден"
    """


class SheetDataException(SheetException):
    """
    Ошибка данных.
    """

class XLSReportWriteException(Exception):
    """
    Ошибка вывода в отчетах XLS
    """

class XLSXReportWriteException(Exception):
    """
    Ошибка вывода в отчетах XLSX
    """

class WrongDocumentType(Exception):
    """
    Ошибка формата документа
    """
    # Например, в Word-документе не предусмотрена генерация таблиц
