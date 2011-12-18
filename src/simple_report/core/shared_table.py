#coding: utf-8
import re
from lxml.etree import _Element, Element, SubElement

__author__ = 'prefer'


class SharedStringsTable(object):
    """
    """

    FIND_PARAMS = re.compile('#\w+#')

    def __init__(self, root):
        """
        """
        # Ключами являются строки, индексами значения
        self.new_elements_dict = {}
        self.new_elements_list = []

        assert isinstance(root, _Element)
        assert 'count' in root.attrib
        assert 'uniqueCount' in root.attrib
        assert  hasattr(root, 'nsmap')

        self.nsmap = root.nsmap
        self.uniq_elements = self.count = 0

        self.elements = [t.text for si in root for t in si]

    def get_new_index(self, index):
        """
        """
        i = int(index)

        value = self.elements[i]

        if value in self.new_elements_list:
            self.uniq_elements += 1
            return str(self.new_elements_dict[value])
        else:
            self.new_elements_list.append(value)
            len_list = len(self.new_elements_list)
            self.new_elements_dict[value] = len_list - 1
            return str(len_list - 1)

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

    def set_params(self, indexes, params):
        """
        """
        if indexes:
            for i in indexes:
                value = self.new_elements_list[i]
                parameters = self.FIND_PARAMS.findall(value)
                for param in parameters:
                    text = params.get(param[1:-1])
                    if text is not None:
                        self.new_elements_list[i] = value.replace(param, unicode(text))
                    else:
                        self.new_elements_list[i] = ''
