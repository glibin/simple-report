#coding: utf-8

from lxml.etree import tostring
from simple_report.core import XML_DEFINITION

from simple_report.core.xml_wrap import ReletionOpenXMLFile, CommonProperties

__author__ = 'prefer'


class Wordprocessing(ReletionOpenXMLFile):
    """

    """
    NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # Узел контекста document
    # .// рекурсивно спускаемся к потомкам в поисках <ns:p><ns:r><ns:t></ns:t></ns:r></ns:p>
    XPATH_QUERY = './/{0}:p/{0}:r/{0}:t'

    def __init__(self, tags, *args, **kwargs):
        super(Wordprocessing, self).__init__(*args, **kwargs)

        self.tags = tags


    def build(self):
        """

        """
        with open(self.file_path, 'w') as f:
            f.write(XML_DEFINITION + tostring(self._root))

    def set_params(self, params):
        """
        """

        text_nodes = self._root.xpath(self.XPATH_QUERY.format('w'), namespaces={'w': self.NS_W})

        for node in text_nodes:
            for key_param, value in params.items():
                if key_param in node.text:
                    if len(node.text) > 0 and node.text[0] == '#' and node.text[-1] == '#':
                        node.text = node.text.replace('#%s#' % key_param, unicode(value))
                    else:
                        node.text = node.text.replace(key_param, unicode(value))

    def get_all_parameters(self):
        """
        """
        text_nodes = self._root.xpath(self.XPATH_QUERY.format('w'), namespaces={'w':self.NS_W})

        for node in text_nodes:
            if len(node.text) > 0 and node.text[0] == '#' and node.text[-1] == '#':
                yield node.text


class CommonPropertiesDOCX(CommonProperties):
    """

    """

    def _get_app_common(self, _id, target):
        """
        """
        return Wordprocessing.create(self.tags, _id, *self._get_path(target))