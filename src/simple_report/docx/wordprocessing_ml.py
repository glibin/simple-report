#coding: utf-8

from lxml.etree import tostring
from simple_report.core import XML_DEFINITION

from simple_report.core.xml_wrap import ReletionOpenXMLFile, CommonProperties

__author__ = 'prefer'


class Wordprocessing(ReletionOpenXMLFile):
    """

    """
    NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    NS_V = 'urn:schemas-microsoft-com:vml'

    XPATH_TEXT = '/{0}:document/{0}:body/{0}:p/{0}:r/{0}:t'
    XPATH_TABLE = '/{0}:document/{0}:body/{0}:tbl/{0}:tr/{0}:tc/{0}:p/{0}:r/{0}:t'
    XPATH_PICTURE = '/{0}:document/{0}:body/{0}:p/{0}:r/{0}:pict/{1}:rect/{1}:textbox/{0}:txbxContent/{0}:p/{0}:r/{0}:t'

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
        text_nodes = self._root.xpath(self.XPATH_TEXT.format('w'), namespaces={'w': self.NS_W})
        text_nodes.extend(self._root.xpath(self.XPATH_TABLE.format('w'), namespaces={'w':self.NS_W}))
        text_nodes.extend(self._root.xpath(self.XPATH_PICTURE.format('w', 'v'), namespaces={'w':self.NS_W,
                                                                                                        'v':self.NS_V}))
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
        text_nodes =self._root.xpath(self.XPATH_TEXT.format('w'), namespaces={'w': self.NS_W})
        text_nodes.extend(self._root.xpath(self.XPATH_TABLE.format('w'), namespaces={'w':self.NS_W}))
        text_nodes.extend(self._root.xpath(self.XPATH_PICTURE.format('w', 'v'), namespaces={'w': self.NS_W,
                                                                                                       'v': self.NS_V}))
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