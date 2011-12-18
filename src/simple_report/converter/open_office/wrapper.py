#coding: utf-8

import os
from exceptions import AttributeError, KeyError, Exception
from simple_report.converter.open_office import settings as st

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.task import ErrorCodeIOException
from com.sun.star.connection import NoConnectException

__author__ = 'prefer'


class OOWrapperException(Exception):
    """

    """

class OOWrapper(object):
    u"""
    Для работы с OpenOffice
    """

    def __init__(self, port=st.DEFAULT_OPENOFFICE_PORT):
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                         localContext)
        try:
            context = resolver.resolve("uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port)
        except NoConnectException:
            raise OOWrapperException("failed to connect to OpenOffice.org on port %s" % port)
        self.desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

    def convert(self, inputFile, _outputExt):
        """
        """
        outputFile, _ext = os.path.splitext(inputFile)
        outputFile = os.path.extsep.join((outputFile, _outputExt))

        inputUrl = self._toFileUrl(inputFile)
        outputUrl = self._toFileUrl(outputFile)

        loadProperties = {"Hidden": True}
        inputExt = self._getFileExt(inputFile)
        if st.IMPORT_FILTER_MAP.has_key(inputExt):
            loadProperties.update(st.IMPORT_FILTER_MAP[inputExt])

        document = self.desktop.loadComponentFromURL(inputUrl, "_blank", 0,
                                                     self._toProperties(loadProperties))
        try:
            document.refresh()
        except AttributeError:
            pass

        family = self._detectFamily(document)
        self._overridePageStyleProperties(document, family)

        outputExt = self._getFileExt(outputFile)
        storeProperties = self._getStoreProperties(document, outputExt)

        try:
            document.storeToURL(outputUrl, self._toProperties(storeProperties))
        finally:
            document.close(True)

        return uno.fileUrlToSystemPath(outputUrl)

    def _overridePageStyleProperties(self, document, family):
        """
        """
        if st.PAGE_STYLE_OVERRIDE_PROPERTIES.has_key(family):
            properties = st.PAGE_STYLE_OVERRIDE_PROPERTIES[family]
            pageStyles = document.getStyleFamilies().getByName('PageStyles')
            for styleName in pageStyles.getElementNames():
                pageStyle = pageStyles.getByName(styleName)
                for name, value in properties.items():
                    pageStyle.setPropertyValue(name, value)

    def _getStoreProperties(self, document, outputExt):
        family = self._detectFamily(document)
        try:
            propertiesByFamily = st.EXPORT_FILTER_MAP[outputExt]
        except KeyError:
            raise OOWrapperException("unknown output format: '%s'" % outputExt)
        try:
            return propertiesByFamily[family]
        except KeyError:
            raise OOWrapperException("unsupported conversion: from '%s' to '%s'" % (family, outputExt) )

    def _detectFamily(self, document):
        if document.supportsService("com.sun.star.text.WebDocument"):
            return st.FAMILY_WEB
        if document.supportsService("com.sun.star.text.GenericTextDocument"):
            # must be TextDocument or GlobalDocument
            return st.FAMILY_TEXT
        if document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            return st.FAMILY_SPREADSHEET
        if document.supportsService("com.sun.star.presentation.PresentationDocument"):
            return st.FAMILY_PRESENTATION
        if document.supportsService("com.sun.star.drawing.DrawingDocument"):
            return st.FAMILY_DRAWING
        raise OOWrapperException("unknown document family: %s" % document)

    def _getFileExt(self, path):
        ext = os.path.splitext(path)[1]
        if ext is not None:
            return ext[1:].lower()

    def _toFileUrl(self, path):
        return uno.systemPathToFileUrl(os.path.abspath(path))

    def _toProperties(self, dict):
        props = []
        for key in dict:
            prop = PropertyValue()
            prop.Name = key
            prop.Value = dict[key]
            props.append(prop)
        return tuple(props)

