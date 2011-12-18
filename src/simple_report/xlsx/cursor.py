#coding: utf-8
__author__ = 'prefer'


class Cursor(object):
    def __init__(self, column=None, row=None, ):
        self._column = column or ('A', 1)
        self._row = row or ('A', 1)

    def __repr__(self):
        return self.__str__()

    def __str__(self):
        return "row: %s - col: %s" % (self._row, self._column)

    @property
    def row(self):
        return self._row

    @row.setter
    def row(self, value):
        self._test_value(value)
        self._row = value

    @property
    def column(self):
        return self._column

    @column.setter
    def column(self, value):
        self._test_value(value)
        self._column = value

    def _test_value(self, value):
        assert isinstance(value, tuple)
        assert len(value) == 2 # Только два элемента: строка и колонка
        assert isinstance(value[0], basestring)
        assert isinstance(value[1], int)
