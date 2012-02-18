#coding: utf-8

from lxml.etree import tostring
from simple_report.core import XML_DEFINITION

from simple_report.core.xml_wrap import ReletionOpenXMLFile, CommonProperties

__author__ = 'prefer'


class Wordprocessing(ReletionOpenXMLFile):
    """

    """
    NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    XPATH_TEXT = '/{0}:document/{0}:body/{0}:p/{0}:r/{0}:t'


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
        for node in text_nodes:
            for key_param, value in params.items():
                if key_param in node.text:
                    node.text = node.text.replace('#%s#' % key_param, value)


class CommonPropertiesDOCX(CommonProperties):
    """

    """

    def _get_app_common(self, _id, target):
        """
        """
        return Wordprocessing.create(self.tags, _id, *self._get_path(target))