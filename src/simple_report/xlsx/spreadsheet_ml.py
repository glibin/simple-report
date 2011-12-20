#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

import os
import re

from lxml.etree import  QName, tostring
from simple_report.core.xml_wrap import OpenXMLFile, ReletionOpenXMLFile
from simple_report.core.shared_table import SharedStringsTable
from simple_report.utils import get_addr_cell
from simple_report.xlsx.cursor import Cursor
from simple_report.xlsx.section import Section, SheetData


class SectionException(Exception):
    """
    """

class SectionNotFoundException(SectionException):
    """
    """

class Comments(OpenXMLFile):
    """
    """
    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

    section_pattern = re.compile(u'[\+|\-]+[A-Za-zА-яА-я0-9_]+')

    def __init__(self, sheet_data, *args, **kwargs):
        super(Comments, self).__init__(*args, **kwargs)

        assert isinstance(sheet_data, SheetData)
        self._sheet_data = sheet_data

        self.sections = {}

        self.comment_list = self._root.find(QName(self.NS, 'commentList'))
        self._create_section()

        # Проверка, правильно ли указаны секции и есть ли конец секции
        self._test_sections()


    def _test_sections(self):
        """
        """
        for section_object in self.sections.values():
            if not section_object.name or not section_object.begin or not section_object.end:
                raise ValueError("Bad section: %s" % section_object)

    def _parse_sections(self, comment_list):
        """
        """

        for comment in comment_list:
            cell = comment.get('ref')
            for text in comment:
                for r in text:
                    for t in r.findall(QName(self.NS, 't')):
                        yield t.text, cell


    def _create_section(self):
        """
        """

        map(self._add_section, self._parse_sections(self.comment_list))

    def _add_section(self, values):
        text = values[0]
        cell = values[1]

        values = self.section_pattern.findall(text)
        addr = get_addr_cell(cell)
        for value in values:
            section_name = self._get_name_section(value)

            # Такой объект должен быть
            if value.startswith('-'):
                # Такой элемент уже должен быть
                if not self.sections.get(section_name):
                    raise SectionException('Start section "%s" not found' % section_name)

                section = self.sections[section_name]

                # Второго конца быть не может
                if section.end:
                    raise SectionException('For section "%s" more than one ending tag' % section_name)

                section.end = addr
            else:
                # Второго начала у секции быть не может
                if self.sections.get(section_name):
                    raise SectionException('For section "%s" more than one beging tag' % section_name)

                self.sections[section_name] = Section(self._sheet_data, section_name, begin=addr)

    def _get_name_section(self, text):
        """
        Возвращает из наименования ++A - название секции
        """
        for i, s in enumerate(text):
            if s.isalpha():
                return text[i:]
        else:
            raise SectionException('Section name bad format "%s"' % text)

    def get_section(self, section_name):
        """
        """
        try:
            section = self.sections[section_name]
        except KeyError:
            raise SectionNotFoundException('Section "%s" not found' % section_name)
        else:
            return section

    def get_sections(self):
        """
        """
        return self.sections.values()

    @classmethod
    def create(cls, cursor, *args, **kwargs):
        return cls(cursor, *args, **kwargs)

    def build(self):
        """
        """
        if len(self.comment_list) > 0:
            self.comment_list.clear()

        with open(self.file_path, 'w') as f:
            f.write(tostring(self._root))


class SharedStrings(OpenXMLFile):
    """
    """
    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def __init__(self, *args, **kwargs):
        super(SharedStrings, self).__init__(*args, **kwargs)
        self.table = SharedStringsTable(self._root)

    def build(self):
        """
        """
        new_root = self.table.to_xml()
        with open(self.file_path, 'w') as f:
            f.write(tostring(new_root))


class App(OpenXMLFile):
    """
    """


class Core(OpenXMLFile):
    """
    """

class WorkbookSheet(ReletionOpenXMLFile):
    """
    """

    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


    def __init__(self, shared_table, name, sheet_id, *args, **kwargs):
        super(WorkbookSheet, self).__init__(*args, **kwargs)
        self.name = name
        self.sheet_id = sheet_id

        # Первый элемент: начало вывода по вертикали, второй по горизонтали
        self.sheet_data = SheetData(self._root,
                                    cursor=Cursor(),
                                    ns=self.NS,
                                    shared_table=shared_table)

        self.drawing, self.comments = self.walk_reletion()


    def walk_reletion(self):
        """
        """
        drawing = comments = None
        if not self._reletion_root is None:
            for elem in self._reletion_root:
                param = (elem.attrib['Id'], elem.attrib['Target'])
                if elem.attrib['Type'] == ReletionTypes.DRAWING:
                    drawing = self._get_drawing(*param)

                elif elem.attrib['Type'] == ReletionTypes.COMMENTS:
                    comments = self._get_comment(*param)

        return drawing, comments


    def _get_comment(self, rel_id, target):
        """
        """
        return Comments.create(self.sheet_data, rel_id, *self._get_path(target))

    def _get_drawing(self, rel_id, target):
        """
        """

    def __str__(self):
        res = [u'Sheet name "{0}":'.format(self.name)]
        if self.comments:
            for section in self.sections:
                res.append(u'\t %s' % section)
        return u'\n'.join(res).encode('utf-8')


    def __repr__(self):
        return self.__str__()

    @property
    def sections(self):
        return self.comments.get_sections()


    def get_section(self, name):
        """
        """
        return self.comments.get_section(name)

    def get_sections(self):
        """
        """
        return self.sections


    def build(self):
        """
        """
        new_root = self.sheet_data.new_sheet()
        with open(self.file_path, 'w') as f:
            f.write(tostring(new_root))

        if self.comments:
            self.comments.build()


class WorkbookStyles(OpenXMLFile):
    """
    """


class Workbook(ReletionOpenXMLFile):
    """
    """
    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def __init__(self, *args, **kwargs):
        """
        """

        super(Workbook, self).__init__(*args, **kwargs)

        self.workbook_style, tmp_sheets, self.shared_strings = self.walk_reletions()
        self.sheets = self.walk(tmp_sheets)

        if self.sheets:
            # По-умолчанию активным считается первый лист
            self._active_sheet = self.sheets[0]
        else:
            raise Exception('Sheets not found')

    def walk_reletions(self):
        workbook_style = shared_strings = None
        sheets = {}
        for elem in self._reletion_root:
            param = (elem.attrib['Id'], elem.attrib['Target'])
            if elem.attrib['Type'] == ReletionTypes.WORKBOOK_STYLE:
                workbook_style = self._get_style(*param)

            elif elem.attrib['Type'] == ReletionTypes.WORKSHEET:
                sheets[elem.attrib['Id']] = elem.attrib['Target']

            elif elem.attrib['Type'] == ReletionTypes.SHARED_STRINGS:
                shared_strings = self._get_shared_strings(*param)

        return workbook_style, sheets, shared_strings

    def walk(self, sheet_reletion):
        """
        """
        sheets = []
        sheets_elem = self._root.find(QName(self.NS, 'sheets'))
        for sheet_elem in sheets_elem:
            name = sheet_elem.attrib['name']
            sheet_id = sheet_elem.attrib['sheetId']
            # state = sheet_elem.attrib['state'] -- В win файле нет такого свойства

            rel_id = sheet_elem.attrib.get(QName(self.NS_R, 'id'))
            target = sheet_reletion[rel_id]
            sheet = self._get_worksheet(rel_id, target, name, sheet_id)
            sheets.append(sheet)

        return sheets

    def _get_style(self, _id, target):
        """
        """

    def _get_worksheet(self, rel_id, target, name, sheet_id):
        worksheet = WorkbookSheet.create(self.shared_table, name, sheet_id, rel_id, *self._get_path(target))
        return worksheet

    def _get_shared_strings(self, _id, target):
        return SharedStrings.create(_id, *self._get_path(target))

    def get_section(self, name):
        """
        """
        return self._active_sheet.get_section(name)

    def get_sections(self):
        """
        """
        return self._active_sheet.get_sections()

    @property
    def active_sheet(self):
        return self._active_sheet

    @active_sheet.setter
    def active_sheet(self, value):
        assert isinstance(value, int)
        self._active_sheet = self.sheets[value]

    def build(self):
        """
        """
        map(lambda x: x.build(), self.sheets)

        self.shared_strings.build()

    @property
    def shared_table(self):
        return self.shared_strings.table


class ReletionTypes(object):
    """
    """
    WORKBOOK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    APP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

    WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    WORKBOOK_STYLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

    SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

    COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"


class CommonProperties(ReletionOpenXMLFile):
    u"""
    Общие настройки
    
    Находит папку _rels и парсит файл .rels в нем
    """

    NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    def __init__(self, *args, **kwargs):
        super(CommonProperties, self).__init__(*args, **kwargs)

        self.core = self.app = self.workbook = None
        self.walk()


    def walk(self):
        """
        """
        for elem in self._root:
            param = (elem.attrib['Id'], elem.attrib['Target'])
            if elem.attrib['Type'] == ReletionTypes.WORKBOOK:
                self.workbook = self._get_workbook(*param)

            elif elem.attrib['Type'] == ReletionTypes.APP:
                self.app = self._get_app(*param)

            elif elem.attrib['Type'] == ReletionTypes.CORE:
                self.core = self._get_core(*param)


    def _get_app(self, _id, target):
        """
        """
        return App.create(_id, *self._get_path(target))

    def _get_core(self, _id, target):
        """
        """
        return Core.create(_id, *self._get_path(target))

    def _get_workbook(self, _id, target):
        """
        """
        return Workbook.create(_id, *self._get_path(target))

    @classmethod
    def create(cls, folder):
        reletion_path = os.path.join(folder, cls.RELETION_FOLDER, cls.RELETION_EXT)
        rel_id = None # Корневой файл связей
        file_name = '' # Не имеет названия, т.к. состоит из расширения .rels
        return cls(rel_id, folder, file_name, reletion_path, )
