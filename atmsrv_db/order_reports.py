import os.path
from datetime import datetime
from email import encoders
from email._header_value_parser import ContentType
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTP

import xlwt

import atmsrv_db.reporter as reporter
from atmsrv_db.gptyp import from_gpdatetime, OrderState
from atmsrv_db.orcl import Orcl
from atmsrv_db.conf import Conf

sqltext_mail = """
select m.a_date_sent, m.a_subject, mb.a_body_part from r_send_mail m
  inner join r_order_mail om on m.a_mail_ref = om.a_mail_ref
  inner join r_mail_body_part mb on mb.a_mail_ref = m.a_mail_ref
where 1 = 1
  and om.a_order_ref = :ref
  and regexp_like(m.a_recipients, 'ncrmailbox')
"""

sqltext_order_list ="""
select o.ref, o.a_number, o.slm_state from r_order o
where 1 = 1
  and o.contr_ref = 1000
  and o.SERVICE_TYPE in (1, 2, 3)
  and (
    o.slm_state in (1, 2, 4, 5, 6, 7, 8, 9, 10, 14, 19, 21) or
    (o.slm_state in (3, 11, 22) and o.DATE_END between 20161210000000 and 20161210235959)
  )
order by o.DATE_REG desc
"""

timestamp_format = '%d.%m.%Y %H:%M'

header = [
    ('Номер заявки', 4000),
    ('Статус', 4500),
    ('Дата отправки письма', 6000),
    ('Заголовок письма', 10000),
    ('Текст письма', 15000),
]

row_beg = 0
col_beg = 0

# addr_from = 'supsoft@diasoft-service.ru'
# addr_to = ['bespalov@diasoft-service.ru']
# addr_cc = ['bespalov@diasoft-service.ru']
# _host = "mail.diasoft-service.ru"
# _port = 587
# _user = 'testorders@diasoft-service.ru'
# _pass = 'UiGmo0DhTu8h'

# order_header = [
#     'Индекс заявки',
#     'Номер заявки',
#     'Статус',
#     'Дата отправки письма',
#     'Заголовок письма',
#     'Текст письма',
# ]


def actual_ncr():

    order = get_actual_ncr_orders()

    timestamp = datetime.now()
    path, filename = order_report(order, timestamp)

    send_report(path, filename, timestamp)


def get_actual_ncr_orders():
    db = Orcl()
    # db = Orcl(uri="prom_ust_atm/121@fast")

    print('Получение списка актуальных заявок...')

    db.sql_exec(sqltext_order_list, {})
    # orders = [[ref, number, OrderState(state_int)] for ref, number, state_int in db.fetchall()]

    orders = []
    for ref, number, state_int in db.fetchall():
        state = OrderState(state_int)

        db.sql_exec(sqltext_mail, {'ref': ref})
        mails = [[from_gpdatetime(date_sent), subject, body.strip()] for date_sent, subject, body in db.fetchall()]
        mails.sort(key=lambda mail: mail[1])            # Сортировка по дате получения письма

        order = [number, state.title]
        if len(mails) > 0:
            order.extend(mails[0])

        orders.append(order)

    return orders

# class Cell(object):
#     def __init__(self, label, style: xlwt.XFStyle()):
#         # self.x = x
#         # self.y = y
#         self.label = label
#         self.style = style
#
#
# class Table(object):
#     def __init__(self, col_beg=0, row_beg=0):
#         self.col_beg = col_beg
#         self.row_beg = row_beg
#         # self.cols = []
#         # self.rows = []
#         self.cells = []
#
#     def row_add(self, cells):
#         self.cells.append([cell for cell in cells])
#
#     def get_header(self):
#         return self.cells[0]
#
#     # def add_cell(self, cell: Cell):
#     #     # col = self.cols.get(cell.x, [])
#     #     # col.append(cell)
#     #     # row = self.rows.get(cell.y, [])
#     #     # row.append(cell)
#     #     self.cells[(cell.x, cell.y)] = cell


# def table_cre(sheet: xlwt.Worksheet, orders):
#
#
#     fields = orders[0].__dict__.values()
#
#     table.row_add([Cell(field.title, style_title) for field in fields if isinstance(field, Field)])
#     for order in orders:
#         cels = [field.vals for field in orders[0].__dict__.values]
#         mail = order.mails[0]
#         cels.extend([[field.vals for field in mail.__dict__.values]])
#         table.row_add()



# class StyleDate(xlwt.XFStyle):
#     def __init__(self):
#         super().__init__()
#         self.num_format_str = 'DD.MM.YYYY HH:MM:SS'


# class OrderExcel(object):
#     header = [
#         ('Номер заявки', 4000),
#         ('Статус', 4500),
#         ('Дата отправки письма', 7000),
#         ('Заголовок письма', 10000),
#         ('Текст письма', 15000),
#     ]
#
#     def __init__(self, row_beg, col_beg):
#         super().__init__()
#         # self.sheet = sheet
#         self.row_beg = row_beg
#         self.col_beg = col_beg
#
#     def header_write(self, sheet: Worksheet):
#         style = StyleTitle()
#
#         for i, field in enumerate(self.__class__.header):
#             row = self.row_beg
#             col = self.col_beg + i
#             title = field[0]
#             width = field[1]
#
#             sheet.write(row, col, title, style)
#             sheet.col(col).width = width

# titles = ['Индекс заявки', 'Номер заявки', 'Статус', 'Дата отправки письма', 'Заголовок письма', 'Текст письма']
# width_list = [4000, 4500, 7000, 10000, 15000]


def order_report(cells, timestamp):

    print('Формирование отчета...')

    wb = xlwt.Workbook()
    sheet = wb.add_sheet('Orders')

    reporter.header_write(sheet, row_beg, col_beg, header)

    for row, row_cells in enumerate(cells, row_beg + 1):
        reporter.row_write(sheet, row, col_beg, row_cells)

    filename = 'orders_{}.xls'.format(timestamp.strftime('%y%m%d-%H'))
    path = os.path.join('data', filename)
    wb.save(path)

    return path, filename

class SendReporConf(Conf):
    _file_name = 'send_report.cfg'

def send_report(path, filename, timestamp):

    # for i, order in list(enumerate(order_list, 1)):
    #
    #     sheet.write(i, 0, order['a_number'])
    #     sheet.write(i, 1, OrderState(order['slm_state']).title)
    #     mail_list = order['mail_list']
    #     if len(mail_list) > 0:
    #         last_mail = mail_list[-1]
    #         sheet.write(i, 2, from_gpdatetime(last_mail['a_date_sent']), style_date)
    #         sheet.write(i, 3, last_mail['a_subject'])
    #         sheet.write(i, 4, last_mail['a_body_part'].strip())
    #
    # filename = 'orders_{}.xls'.format(timestamp.strftime('%y%m%d-%H'))
    # path = os.path.join('data', filename)
    # wb.save(path)

    conf = SendReporConf()
    addr_from = conf.get('addr_from')
    addr_to = conf.get('addr_to')
    addr_cc = conf.get('addr_cc')
    _host = conf.get('host')
    _port = conf.get('port')
    _user = conf.get('user')
    _pass = conf.get('pass')

    print('Отправка отчета...')
    msg = MIMEMultipart('mixed')
    msg['Subject'] = 'Актуальные заявки {}'.format(timestamp.strftime(timestamp_format))
    msg['From'] = addr_from
    msg['To'] = ', '.join(addr_to)
    msg['cc'] = ', '.join(addr_cc)

    part = MIMEText("""Добрый день.

Во вложении находится список актуальных заявок.

Отчет сформирован {}""".format(timestamp.strftime(timestamp_format)))
    msg.attach(part)

    with open(path, 'rb') as fp:
        part = MIMEBase('application', 'vnd.ms-excel')
        part.set_payload(fp.read())
        encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(part)

    with SMTP(host=_host, port=_port) as smtp:
        smtp.starttls()
        smtp.login(_user, _pass)
        smtp.sendmail(addr_from, addr_to + addr_cc, msg.as_string())

    print('Отчет отправлен.')


# def test_objects():
#
#     # order = Order(100, 'W11111', OrderState(0))
#     # order = Order()
#
#     # print([field.title for field in order.__dict__.values()])
#
#     # order.ref.val = 200
#     # order.ref.val = 300
#     # print(order.ref.val)
#     # print(order.ref.title)
#     # print(order.ref)
#     # order.number.val = 'W22222'
#     # order.state.val = OrderState(1)
#     # print(order.number)
#     #
#     # order2 = Order(500, 'W22222', OrderState(1))
#     # # order2 = Order(500)
#     # order2.number.val = 'W1111'
#     # order2.state = OrderState(0)
#     #
#     # print(order.number.val)
#     # print(order.ref.val)
#     #
#     # print(order2.number.val)
#     # print(order2.ref.val)

# def sent_mail():
#     print('Отправка письма...')
#     msg = MIMEText('Тестовое письмо')
#     msg['Subject'] = 'Тестовое письмо'
#     msg['From'] = 'sd-exch1 <sd-exch1@sberbank.ru>'
#     msg['To'] = 'bespalov@diasoft-service.ru'
#     # msg['cc'] = ', '.join(addr_cc)
#
#     # with open(path, 'rb') as fp:
#     #     part = MIMEBase('application', 'vnd.ms-excel')
#     #     part.set_payload(fp.read())
#     #     encoders.encode_base64(part)
#     # part.add_header('Content-Disposition', 'attachment', filename=filename)
#     # msg.attach(part)
#
#     with SMTP(host=_host, port=_port) as smtp:
#         smtp.starttls()
#         smtp.login(_user, _pass)
#         smtp.sendmail(addr_from, addr_to + addr_cc, msg.as_string())
#
#     print('Письмо отправлено.')
