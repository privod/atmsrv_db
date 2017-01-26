from datetime import datetime

import xlwt
from xlwt import Worksheet


class Reporter(object):
    def __init__(self):
        pass

class EcxelReporter(Reporter)
    def __init__(self):
        super().__init__()



class FontTitle(xlwt.Font):
    def __init__(self):
        super().__init__()
        self.bold = True


class StyleTitle(xlwt.XFStyle):
    def __init__(self):
        super().__init__()
        self.font = FontTitle()

def row_write(sheet: Worksheet, row_beg, col_beg, cells, style=xlwt.XFStyle()):
    for i, cell in enumerate(cells):
        row = row_beg
        col = col_beg + i
        if isinstance(cell, datetime):
            style.num_format_str = 'DD.MM.YYYY HH:MM:SS'
        sheet.write(row, col, cell, style)
