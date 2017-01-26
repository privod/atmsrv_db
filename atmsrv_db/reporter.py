from datetime import datetime

import xlwt
from xlwt import Worksheet

# class Reporter(object):
#     def __init__(self):
#         pass
#
# class EcxelReporter(Reporter)
#     def __init__(self):
#         super().__init__()

class FontTitle(xlwt.Font):
    def __init__(self):
        super().__init__()
        self.bold = True


class StyleTitle(xlwt.XFStyle):
    def __init__(self):
        super().__init__()
        self.font = FontTitle()


# class StyleDate(xlwt.XFStyle):
#     def __init__(self):
#         super().__init__()
#         self.num_format_str = 'DD.MM.YYYY HH:MM:SS'

_style = xlwt.XFStyle()

def header_write(sheet: Worksheet, row, col_beg, header):
    style = StyleTitle()

    for i, field in enumerate(header):
        col = col_beg + i
        title, width = field

        sheet.write(row, col, title, style)
        sheet.col(col).width = width


def row_write(sheet: Worksheet, row, col_beg, cells):
    for i, cell in enumerate(cells):
        col = col_beg + i
        if isinstance(cell, datetime):
            _style.num_format_str = 'DD.MM.YYYY HH:MM:SS'
        sheet.write(row, col, cell, _style)
