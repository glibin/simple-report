#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''
from datetime import datetime

import os
import sys
import shutil
from tempfile import gettempdir
import uuid
import zipfile
from collections import namedtuple

class FileException(Exception):
    pass


class ZipFileAdapter(object):
    """
    Из-за того, что в версии начиная с 2.7 модуль zipfile умеет работать с менеджером контекста, а вот ниже версии - не умеет.
     Будем поддерживать версию 2.6
    """

    def __init__(self, *args, **kwargs):
        self.z_file = zipfile.ZipFile(*args, **kwargs)

        self.old = False
        # Проверка версии в питоне
        if sys.version_info < (2, 7):
            self.old = True

    def __enter__(self):
        if self.old:
            return self.z_file
        else:
            return self.z_file.__enter__()

    def __exit__(self, *args, **kwargs):
        if self.old:
            return self.z_file.close()
        else:
            return self.z_file.__exit__(*args, **kwargs)


class ZipProxy(object):
    u"""
    Распаковка/упаковка Open XML
    """

    @classmethod
    def _extract(cls, src_file_path, dst_files_path):
        u"""
        Распоковывает zip файл
        """
        with ZipFileAdapter(src_file_path) as zip_file:
            zip_file.extractall(dst_files_path)

    @classmethod
    def _pack(cls, dst_file_path, src_files_path):
        u"""
        Запаковывает zip файл
        """
        with ZipFileAdapter(dst_file_path, 'w') as zip_file:
            for root, _, file_names in os.walk(src_files_path):
                for file_name in file_names:
                    # Абсолютный путь до файла
                    abs_path = os.path.join(root, file_name)

                    # Путь до директории
                    dir_path = abs_path[len(src_files_path) + len(os.sep):]

                    zip_file.write(abs_path, dir_path, compress_type=zipfile.ZIP_DEFLATED)

    @classmethod
    def extract(cls, src_file):
        assert isinstance(src_file, FileProxy)

        # Распаковываем zip архив во временную папку
        extract_folder = os.path.join(gettempdir(),
            '_'.join([str(uuid.uuid4())[:8], src_file.get_file_name()]))

        cls._extract(src_file.get_path(), extract_folder)

        return extract_folder

    @classmethod
    def pack(cls, dst_file, extract_folder):
        u"""
        Запаковать в файл
        """
        assert isinstance(dst_file, FileProxy)
        cls._pack(dst_file.get_path(), extract_folder)
        shutil.rmtree(extract_folder)


class FileProxy(object):
    u"""
    """

    def __init__(self, file_like_object, new_file=False):
        if isinstance(file_like_object, FileProxy):
            file_like_object = file_like_object.file

        self.is_file_like_object = False
        if hasattr(file_like_object, 'read'):
            raise FileException("File like object temporarily not supported.")

        if not os.path.exists(file_like_object) and not new_file:
            raise FileException('File "%s" not found.' % file_like_object)

        if not os.path.isfile(file_like_object) and not new_file:
            raise FileException('"%s" is not file' % file_like_object)

        self.file = file_like_object

    def get_path(self):
        u"""
        Возвращает путь до файла
        """
        return self.file

    def get_file_like_object(self):
        u"""
        Возвращает открытый файл
        """
        return self.file

    def get_file_name(self):
        u"""
        Возвращает только имя файла
        """
        if self.is_file_like_object:
            file_name = self.file.name
        else:
            file_name = self.file

        return os.path.split(file_name)[-1]


class ColumnHelper(object):
    """
    Существенно облегчает работу с ячейками

    По данными ('ALC'), ('AVB') возврашает итератор по диапазону ячеек
    ('ALC', 'ALD', 'ALE', ..., 'AVA', 'AVB')

    Так же умеет работать, если переданы данные: ('A'), ('ALP')

    """

    @classmethod
    def number_to_column(cls, n):
        """
        """
        return ~n and cls.number_to_column(n / 26 - 1) + chr(65 + n % 26) or ''

    @classmethod
    def column_to_number(cls, index):
        """
        """
        s = 0
        pow_ = 1
        for letter in index[::-1]:
            d = int(letter, 36) - 9
            s += pow_ * d
            pow_ *= 26
            # excel starts column numeration from 1
        return s - 1


    @classmethod
    def get_range(cls, begin, end):
        """
        """
        for i in xrange(cls.column_to_number(begin), cls.column_to_number(end) + 1):
            yield cls.number_to_column(i)

    @classmethod
    def add(cls, column, i):
        """
        Добавляет к колонке column i колонок
        """
        return cls.number_to_column(cls.column_to_number(column) + i)

    @classmethod
    def difference(cls, col1, col2):
        """
        Разница между двумя колонками
        """
        return cls.column_to_number(col1) - cls.column_to_number(col2)


def get_addr_cell(text):
    """
    Возвращает адрес ячейки
    То есть из представления 'AZ12' выдается ('AZ', 12)
    """
    for i, s in enumerate(text):
        if s.isdigit():
            return text[:i], int(text[i:])
    else:
        raise ValueError('Addr cell is bad format value "%s"' % text)


def date_to_float(date):
    """
    Конвертирует дату в число с плавающей точкой относительно 1900 года

    Значение даты OLE-автоматизации реализовано как число с плавающей запятой,
    равное количеству дней, прошедших с полуночи 30 декабря 1899 г.
    Например, полночь 31 декабря 1899 г. представляется числом 1,0;
    06:00 1 января 1900 г. — числом 2,25;
    полночь 29 декабря 1899 г. — числом 1,0
    и 06:00 29 декабря 1899 г. — числом 1,25.

    http://msdn.microsoft.com/ru-ru/library/system.datetime.tooadate%28v=vs.90%29.aspx

    ps: Спасибо Вадиму, который предотвратил распространение кастылей,
    типо кастомного форматирования параметров
    """
    assert isinstance(date, datetime)

    date_1900 = datetime(1899, 12, 30)

    days = abs((date - date_1900).days)

    hours = date.time().hour - date_1900.time().hour

    minute = date.time().minute - date_1900.time().minute

    return days + (hours + minute / 60.0) / 24.0


FormulaWriteExcel = namedtuple('FormulaWriteExcel',
                                ['formula_id', 'excel_function', 'ranged'],
                                verbose=False)
