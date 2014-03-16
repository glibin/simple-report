=====================
Примеры использования
=====================

Отчеты в форматах `xslx` и `xls`
-------------------------------

Простой пример

    from simple_report.report import SpreadsheetReport
    from simple_report.utils import FormulaWriteExcel
    from simple_report.xls.document import DocumentXLS
    from simple_report.xlsx.document import DocumentXLSX

    # Указываем путь до шаблона
    path_to_template = '/home/user/report_template.xlsx'

    # В самом простом варианте ничего, кроме шаблона, указывать не надо
    report = SpreadsheetReport(
        path_to_template
    )
    header_section = report.get_section('header')
    header_section.flush({
        'author': u'BARS Group',
        'today': datetime.date.today()
    })


Отчеты в формате `docx`
-----------------------



