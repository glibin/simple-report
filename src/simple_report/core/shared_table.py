#coding: utf-8
from lxml.etree import _Element, Element, SubElement

__author__ = 'prefer'


class SharedStringsTable(object):
    """
    """

    def __init__(self, root):
        """
        """
        # Ключами являются строки, индексами значения
        self.new_elements_dict = {}
        self.new_elements_list = []

        assert isinstance(root, _Element)
        #assert 'count' in root.attrib # TODO: Разобраться почему нет в виндовых шаблонах
        #assert 'uniqueCount' in root.attrib # TODO: Разобраться почему нет в виндовых шаблонах
        assert  hasattr(root, 'nsmap')

        self.nsmap = root.nsmap
        self.uniq_elements = self.count = 0

        self.elements = [''.join(t.text or '' for t in si) for si in root]

    def get_new_index(self, index):
        """
        """
        i = int(index)

        value = self.elements[i]

        if value in self.new_elements_list:
            return str(self.new_elements_dict[value])
        else:
            self.uniq_elements += 1
            self.new_elements_list.append(value)
            len_list = len(self.new_elements_list)
            self.new_elements_dict[value] = len_list - 1
            return str(len_list - 1)

    def get_value(self, index):
        """
        Возвращает значение ячейки по индексу
        @type index: int
        """
        return self.elements[index]

    def to_xml(self):
        u"""
        Переводит таблицу в xml
        Возвращает корневой узел xml
        """
        root = Element('sst', {'count': str(len(self.new_elements_list)),
                               'uniqueCount': str(self.uniq_elements)},
            nsmap=self.nsmap)

        for elem in self.new_elements_list:
            si = SubElement(root, 'si')
            t = SubElement(si, 't')
            t.text = unicode(elem)

        return root
