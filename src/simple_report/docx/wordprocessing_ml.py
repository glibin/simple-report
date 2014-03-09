#!coding:utf-8
import copy
from lxml.etree import tostring
from simple_report.core import XML_DEFINITION
from simple_report.core.exception import (
    SectionNotFoundException, SectionException)
from simple_report.core.xml_wrap import (
    ReletionOpenXMLFile, CommonProperties)


__author__ = 'prefer'


class Wordprocessing(ReletionOpenXMLFile):
    """
    Основной файл формата DOCX
    """
    NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # Узел контекста document
    # .// рекурсивно спускаемся к потомкам в поисках
    # <ns:p><ns:r><ns:t></ns:t></ns:r></ns:p>
    XPATH_QUERY = './/{0}:p/{0}:r/{0}:t'
    TABLES_QUERY = './/{0}:tbl'
    TABLE_TEXT_NODE_QUERY = './/{0}:tc/{0}:p/{0}:r/{0}:t'

    def __init__(self, tags, *args, **kwargs):
        super(Wordprocessing, self).__init__(*args, **kwargs)

        self.tags = tags
        self.table_sections = {}

    def build(self):
        """
        Сборка файла
        """
        with open(self.file_path, 'w') as f:
            f.write(XML_DEFINITION + tostring(self._root))

    def set_params(self, params):
        """
         Подстановка параметров
        :param params: параметры подстановки
        @result: None
        """
        #
        self.merge_same_nodes()
        text_nodes = self._root.xpath(
            self.XPATH_QUERY.format('w'), namespaces={'w': self.NS_W}
        )
        self._set_params(text_nodes, params)

    def get_signature(self, node):
        signature = []
        for subnode in list(node):
            if subnode.tag != '{%s}lang' % self.NS_W:
                signature.append((subnode.tag, sorted(subnode.items())))
        return signature

    def merge_same_nodes(self):
        """
         Слияние одинаковых нод - нужно, т.к. редакторы DOCX
        привносят свои специфичные изменения, которые нам не нужны
        @result:
        """
        paragraphs = list(self._root.xpath(
            './/{0}:p'.format('w'), namespaces={'w': self.NS_W}
        ))

        t_tag = '{%s}t' % self.NS_W
        r_tag = '{%s}r' % self.NS_W
        rpr_tag = '{%s}rPr' % self.NS_W
        tab_tag = '{%s}tab' % self.NS_W
        for paragraph in paragraphs:
            par_nodes = list(paragraph)
            old_signature = None
            signature = None
            old_node = None

            for par_node in par_nodes:
                if par_node.tag != r_tag:
                    old_node = None
                    continue
                for node in list(par_node):
                    if node.tag == rpr_tag:
                        old_signature = signature
                        signature = self.get_signature(node)
                    elif node.tag == tab_tag:
                        old_node = None
                    elif node.tag == t_tag:
                        if old_node is not None and old_signature == signature:
                            # delete r node
                            old_node[1].text = old_node[1].text + node.text
                            #old_node = (par_node, node)
                            paragraph.remove(par_node)
                        else:
                            old_node = (par_node, node)

    @classmethod
    def _set_params(cls, text_nodes, params):

        def sorting_key((k, v)):
            if not isinstance(k, basestring):
                return 1
            return -len(k)

        for node in text_nodes:
            for key_param, value in sorted(params.items(), key=sorting_key):
                if key_param in node.text:
                    # if len(node.text) > 0 and node.text[0] == '#' and
                    # node.text[-1] == '#':
                    text_to_replace = '#%s#' % key_param
                    if text_to_replace in node.text:
                        node.text = node.text.replace(
                            text_to_replace, unicode(value))
                    else:
                        node.text = node.text.replace(
                            key_param, unicode(value))

    def get_all_parameters(self):
        """
        Получение всех параметров
        """
        text_nodes = self._root.xpath(
            self.XPATH_QUERY.format('w'), namespaces={'w': self.NS_W})

        for node in text_nodes:
            if (
                len(node.text) > 0 and
                node.text[0] == '#' and
                node.text[-1] == '#'
            ):
                yield node.text

    def get_tables(self):
        """
        Получаем таблицы в DOCX
        """

        return self._root.findall(
            self.TABLES_QUERY.format('w'), namespaces={'w': self.NS_W}
        )

    def set_docx_table_sections(self):
        """
         установка секций таблиц в документах DOCX
        @result: None
        """
        tables = self.get_tables()
        for table in tables:
            text_nodes = table.findall(
                '{0}:tr'.format('w'),
                namespaces={'w': self.NS_W}
            )
            section = Section(table)
            section_name = None
            rows_to_delete = []
            for row_node in text_nodes:
                col_nodes = row_node.findall(
                    self.TABLE_TEXT_NODE_QUERY.format('w'),
                    namespaces={'w': self.NS_W}
                )
                if not col_nodes:
                    continue
                col_node = col_nodes[0]
                if col_node.text and col_node.text[:2] == '#!':
                    section_name = col_node.text[2:]
                    if section_name in self.table_sections:
                        raise SectionException(
                            ('Section named {0} has been found '
                             'more than 1 time in docx table').format(
                                section_name
                            )
                        )
                    rows_to_delete.append(row_node)
                    # del text_nodes[text_nodes.index(row_node)]
                elif col_node.text and col_node.text[-2:] == '!#':
                    self.table_sections[section_name] = section
                    section_name = None
                    section = Section(table)
                    rows_to_delete.append(row_node)
                elif section_name:
                    section.append(copy.copy(row_node))
                    # del text_nodes[text_nodes.index(row_node)]
                    rows_to_delete.append(row_node)
            for row_node in rows_to_delete:
                row_node.getparent().remove(row_node)

    def get_section(self, section_name):
        """
         Получение секции таблицы в документе DOCX по имени
        :param section_name: имя секции
        @result: секция
        """
        if not self.table_sections:
            self.set_docx_table_sections()
        section = self.table_sections.get(section_name)
        if section is None:
            raise SectionNotFoundException(
                'Section named {0} has not been found'.format(section_name)
            )
        return section


class CommonPropertiesDOCX(CommonProperties):

    """

    """

    def _get_app_common(self, _id, target):
        """
        """
        return Wordprocessing.create(self.tags, _id, *self._get_path(target))


class Section(object):
    """
    Секция таблицы docx документа. Поддерживает ограниченное число операций
    В частности, строчки таблицы выводятся полностью, т.е. минимальной
    единицей секции является строка таблицы
    """

    def __init__(self, table):
        self._content = []
        self.table = table

    def append(self, table_row):
        self._content.append(table_row)

    def flush(self, params):
        for row in self._content:
            new_row = copy.copy(row)
            text_nodes = new_row.findall(
                './/{0}:tc/{0}:p/{0}:r/{0}:t'.format('w'),
                namespaces={'w': Wordprocessing.NS_W}
            )
            Wordprocessing._set_params(text_nodes, params)

            self.table.append(new_row)
