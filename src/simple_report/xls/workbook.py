#coding: utf-8

import xlrd
from uuid import uuid4
from xlrd.sheet import Sheet
from xlutils.filter import XLWTWriter
from simple_report.core.exception import SheetNotFoundException, SectionNotFoundException
from simple_report.xls.section import Section
from simple_report.converter.abstract import FileConverter


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
                raise SectionNotFoundException('Section named %s has not been found' % name)
            self.sections[name] = Section(self, name, begin, end, self.writer)
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
            raise SheetNotFoundException('Sheets not found')

        for k, v in kwargs.items():
            setattr(self, k, v)

        self.configure_writer()

    def configure_writer(self):
        """
        Настройка конфигураций writer-а
        """

        if hasattr(self, 'fit_num_pages'):
            self.xlwt_writer.wtsheet.fit_num_pages = self.fit_num_pages

        if hasattr(self, 'portrait_orientation'):
            self.xlwt_writer.wtsheet.portrait = self.portrait_orientation

        if hasattr(self, 'fit_width_to_pages'):
            self.xlwt_writer.wtsheet.fit_width_to_pages = self.fit_width_to_pages

        if hasattr(self, 'fit_height_to_pages'):
            self.xlwt_writer.wtsheet.fit_height_to_pages = self.fit_height_to_pages
        else:
            # По-умолчанию, указываем значение 0, для того, чтобы не запихивать огромный отчет на одну страницу.
            self.xlwt_writer.wtsheet.fit_height_to_pages = 0

        if hasattr(self, 'header_str'):
            self.xlwt_writer.wtsheet.header_str = self.header_str
        else:
            #
            self.xlwt_writer.wtsheet.header_str = (u'')

        if hasattr(self, 'footer_str'):
            self.xlwt_writer.wtsheet.footer_str = self.footer_str
        else:
            #
            self.xlwt_writer.wtsheet.footer_str = (u'&P из %s') % self.write_sheet_count()

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
        try:
            self._active_sheet = self.sheets[value]
        except IndexError:
            raise SheetNotFoundException('Sheet not found')

        try:
            self.xlwt_writer.sheet(self._active_sheet.sheet, self.get_sheet_name())
            # Проставляем номер листа в выходном документе.
            # Для этого берём номер соответствующего узла из шаблона и прибавляем 1, т.к.
            # нумерация с нуля.
            self.xlwt_writer.wtsheet.start_page_number = self.xlwt_writer.rdsheet.number + 1
            # Настраиваем writer для работы с новым листом.
            self.configure_writer()
        except ValueError:
            return

    def _sheet_list(self):
        all_sheets = self.workbook._sheet_list

        sheet_list = []
        for sheet in all_sheets:
            sheet_list.append(WorkbookSheet(sheet, self.xlwt_writer))

        return sheet_list

    def get_sheet_name(self):
        return self.active_sheet.get_name()

    def build(self, dest_file):
        """
        """

        dest_file_name = dest_file.file

        self.xlwt_writer.finish()

        # Получаем формат файла
        dest_file_format = dest_file_name.split('.')[-1]

        # Если не xls, то проставляем явно.
        if dest_file_format != FileConverter.XLS:
            dest_file_name = '%s.%s' % (dest_file_name, FileConverter.XLS)

        # self.xlwt_writer.output имеет вид
        # [('выходной файл1', Workbook1), ('выходной файл2', Workbook2), ... ]
        # Для данного Workbook выбираем первый кортеж
        ouput_file, workbook = self.xlwt_writer.output[0]
        workbook.save(dest_file_name)

    def write_sheet_count(self):
        """
        Подсчитываем количество листов в которых есть секции (другими словами в них будет производиться запись)
        """

        # Берём листы из шаблона
        read_sheets = self.xlwt_writer.rdbook._sheet_list

        return len([read_sheet for read_sheet in read_sheets if read_sheet.cell_note_map])