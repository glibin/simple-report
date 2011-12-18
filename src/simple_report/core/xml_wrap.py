#coding: utf-8
from abc import ABCMeta
import os
from lxml.etree import parse

__author__ = 'prefer'


class OpenXMLFile(object):
    """

    """
    __metaclass__ = ABCMeta

    NS = None

    def __init__(self, rel_id, folder, file_name, file_path):
        self.reletion_id = rel_id
        self.current_folder = folder
        self.file_name = file_name
        self.file_path = file_path
        self._root = self.from_file(file_path)

    def _get_path(self, target):
        """
            Возвращает относительный путь, название файла, полный путь
            """
        split_path = os.path.split(target)
        relative_path = os.path.join(self.current_folder, *split_path[:-1])
        abs_path = os.path.join(relative_path, split_path[-1])

        return relative_path, split_path[-1], abs_path


    def get_root(self):
        return self._root

    @classmethod
    def from_file(cls, file_path):
        assert file_path
        with open(file_path) as f:
            return parse(f).getroot()


    @classmethod
    def create(cls, *args, **kwargs):
        """
        """
        return cls(*args, **kwargs)


class ReletionOpenXMLFile(OpenXMLFile):
    __metaclass__ = ABCMeta

    RELETION_EXT = '.rels'
    RELETION_FOLDER = '_rels'

    def __init__(self, *args, **kwargs):
        super(ReletionOpenXMLFile, self).__init__(*args, **kwargs)

        assert not self.file_name is None

        rel_path = os.path.join(self.current_folder, self.RELETION_FOLDER, self.file_name + self.RELETION_EXT)

        self._reletion_root = None # Если остальные листы в документе не используются, то стилей для них нет
        if os.path.exists(rel_path):
            self._reletion_root = self.from_file(rel_path)

