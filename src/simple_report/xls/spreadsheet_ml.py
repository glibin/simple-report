#coding: utf-8

import xlrd
from uuid import uuid4
from xlrd.sheet import Sheet
from xlutils.filter import XLWTWriter
from simple_report.xls.section import Section


class WorkbookSheet():

    def __init__(self, sheet, writer):
        assert isinstance(sheet, Sheet), 'sheet must be xlrd.sheet.Sheet instance'

        self.sheet = sheet
        self.writer = writer
        self.wtrowx = 0
        self.sections = {}

    def get_section(self, name):

        assert name, 'Section name is empty'
        if name not in self.sections:
            begin = end = None

            begin_section_text = u''.join(['+', name])
            end_section_text = u''.join(['-', name])

            # Получаем все примечания.
            # sheet.cell_note_map представляет собой словарь
            # { (colx, coly): note_instance }
            notes = self.sheet.cell_note_map

            for (note_coord, note) in notes.items():
                note_text = note.text

                if note_text == begin_section_text:
                    begin = note_coord

                elif note_text == end_section_text:
                    end = note_coord

                elif note_text in u''.join([begin_section_text, end_section_text]):
                    begin = end = note_coord

            if not (begin and end):
                raise Exception('Section named %s has not been found' % name)
            self.sections[name] = Section(begin, end, self, self.writer)
        return self.sections[name]

    def get_sections(self):
        return self.sections

    def get_name(self):
        return self.sheet.name


class Workbook(object):

    def __init__(self, ffile, *args, **kwargs):
        """
        """

        self.workbook = xlrd.open_workbook(ffile.file, formatting_info=True)

        self.xlwt_writer = XLWTWriter()
        self.xlwt_writer.start()
        self.xlwt_writer.workbook(self.workbook, '%s.xls'%uuid4())

        self.sheets = self._sheet_list()

        if self.sheets:
            self._active_sheet = self.sheets[0]
            self.xlwt_writer.sheet(self._active_sheet.sheet, self._active_sheet.sheet.name)
        else:
            raise Exception('Sheets not found')

        for k, v in kwargs.items():
            self.__setattr__(k, v)

    def get_section(self, name):
        return self._active_sheet.get_section(name)

    def get_sections(self):

        workbook_sections = {}

        for sheet in self.sheets:
            workbook_sections.update(sheet.get_sections())

        return workbook_sections

    @property
    def active_sheet(self):
        return self._active_sheet

    @active_sheet.setter
    def active_sheet(self, value):
        assert isinstance(value, int)
        self._active_sheet = self.sheets[value]
        self.xlwt_writer.sheet(self._active_sheet.sheet, self.get_sheet_name())

    def _sheet_list(self):
        all_sheets = self.workbook._sheet_list

        sheet_list = []
        for sheet in all_sheets:
            sheet_list.append(WorkbookSheet(sheet, self.xlwt_writer))

        return sheet_list

    def get_sheet_name(self):
        return self.active_sheet.get_name()

    def show(self, dest_file_name, file_type=None):
        """
        """

        if hasattr(self, 'fit_num_pages'):
            self.xlwt_writer.wtsheet.fit_num_pages = self.fit_num_pages
        if hasattr(self, 'portrait_orientation'):
            self.xlwt_writer.wtsheet.portrait = self.portrait_orientation
        if hasattr(self, 'fit_width_to_pages'):
            self.xlwt_writer.wtsheet.fit_width_to_pages = self.fit_width_to_pages
        if hasattr(self, 'fit_height_to_pages'):
            self.xlwt_writer.wtsheet.fit_height_to_pages = self.fit_height_to_pages
        self.xlwt_writer.finish()

        if file_type and not dest_file_name.endswith(file_type):
            dest_file_name = '%s.%s' % (dest_file_name, file_type)

        # self.xlwt_writer.output имеет вид
        # [('выходной файл1', Workbook1), ('выходной файл2', Workbook2), ... ]
        # Для данного Workbook выбираем первый кортеж
        ouput_file, workbook = self.xlwt_writer.output[0]
        workbook.save(dest_file_name)
