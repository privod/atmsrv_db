import os.path
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.message import MIMEMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTP

import xlwt

from atmsrv_db.gptyp import from_gpdatetime, OrderState
from atmsrv_db.orcl import Orcl

sqltext_mail = """
select m.a_mail_ref, m.a_date_sent, m.a_subject, mb.a_body_part from r_send_mail m
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

addr_from = 'supsoft@diasoft-service.ru'
addr_to = ['bespalov@diasoft-service.ru']
addr_cc = ['bespalov@diasoft-service.ru']
_host = "mail.diasoft-service.ru"
_port = 587
_user = 'testorders@diasoft-service.ru'
_pass = 'UiGmo0DhTu8h'


# class FieldInfo(object):
#     def __init__(self, title=None):
#         self.title = title

# class _DomainMetaclass(type):
#     def __init__(cls, name, bases, attributes):
#         super().__init__(name, bases, attributes)
#         print(cls)
#         print(name)
#         # print(attributes)
#         # print(attributes.get('fields'))
#         for field in attributes.get('fields', []):
#             print(field)


class Field(object):
    title = None

    def __init__(self, val=None):
        self.val = val


class FieldRef(Field):
    title = 'Индекс'


class FieldDateSent(Field):
    title = 'Дата отправки письма'


class FieldSubject(Field):
    title = 'Заголовок письма'


class FieldBody(Field):
    title = 'Текст письма'


class FieldNumber(Field):
    title = 'Номер заявки'


class FieldState(Field):
    title = 'Статус'


# class Model(object, metaclass=_DomainMetaclass):
# # class Model(object):
#
#     def __init__(self):
#         # print(self.__class__)
#         # print(self.__class__.__mro__)
#         for cls in self.__class__.__mro__:
#             for name, title in cls.__dict__.get('fields', []):
#                 # print(name, title)
#
#                 self.__dict__[name] = None
#                 self.__class__.__dict__['{}_title'.format(name)] = title
#
#         # print(dir(self))
#         # print(getattr(self, 'fields'))
#         pass


class Domain(object):
    def __init__(self, ref):
        self.ref = FieldRef(ref)


class Mail(Domain):
    def __init__(self, ref, date_sent, subject, body):
        super().__init__(ref)
        self.date_sent = FieldDateSent(date_sent)
        self.subject = FieldSubject(subject)
        self.body = FieldBody(body)


class Order(Domain):
    def __init__(self, ref, number, state):
        super().__init__(ref)
        self.number = FieldNumber(number)
        self.state = FieldState(state)


def actual_ncr():
    db = Orcl()
    # db = Orcl(uri="prom_ust_atm/121@fast")

    print('Получение списка актуальных заявок...')
    # order_list = db.exec_in_dict(sqltext_order_list, {})
    #
    # for order in order_list:
    #     mail_list = db.exec_in_dict(sqltext_mail, {'order_number': order['a_number']})
    #     order['mail_list'] = sorted(mail_list, key=lambda send: send['a_date_sent'])

    db.sql_exec(sqltext_order_list, {})
    orders = [(ref, number, OrderState(state_int)) for ref, number, state_int in db.fetchall()]

def actual_ncr_objects():
    db = Orcl()

    db.sql_exec(sqltext_order_list, {})
    orders = [Order(ref, number, OrderState(state_int)) for ref, number, state_int in db.fetchall()]

    for order in orders:
        db.sql_exec(sqltext_mail, {'ref': order.ref.val})
        mails = [Mail(ref, date_sent, subject, body) for ref, date_sent, subject, body in db.fetchall()]
        order.mails = sorted(mails, key=lambda mail: mail.date_sent.val)

    return orders


class Cell(object):
    def __init__(self, label, style: xlwt.XFStyle()):
        # self.x = x
        # self.y = y
        self.label = label
        self.style = style


class Table(object):
    def __init__(self, col_beg=0, row_beg=0):
        self.col_beg = col_beg
        self.row_beg = row_beg
        # self.cols = []
        # self.rows = []
        self.cells = []

    def row_add(self, cells):
        self.cells.append([cell for cell in cells])

    def get_header(self):
        return self.cells[0]

    # def add_cell(self, cell: Cell):
    #     # col = self.cols.get(cell.x, [])
    #     # col.append(cell)
    #     # row = self.rows.get(cell.y, [])
    #     # row.append(cell)
    #     self.cells[(cell.x, cell.y)] = cell


def table_cre(sheet: xlwt.Worksheet, orders):

    font_title = xlwt.Font()
    font_title.bold = True

    style_title = xlwt.XFStyle()
    style_title.font = font_title
    style_date = xlwt.XFStyle()
    style_date.num_format_str = 'DD.MM.YYYY HH:MM:SS'

    table = Table()

    fields = orders[0].__dict__.values()

    table.row_add([Cell(field.title, style_title) for field in fields if isinstance(field, Field)])
    for order in orders:
        cels = [field.vals for field in orders[0].__dict__.values]
        mail = order.mails[0]
        cels.extend([[field.vals for field in mail.__dict__.values]])
        table.row_add()


def report(table):
    print('Формирование отчета...')
    timestamp = datetime.now()

    wb = xlwt.Workbook()
    sheet = wb.add_sheet('Orders')

    font_title = xlwt.Font()
    font_title.bold = True

    style_title = xlwt.XFStyle()
    style_title.font = font_title
    style_date = xlwt.XFStyle()
    style_date.num_format_str = 'DD.MM.YYYY HH:MM:SS'

    sheet.write(0, 0, 'Номер заявки', style_title)
    sheet.col(0).width = 4000
    sheet.write(0, 1, 'Статус', style_title)
    sheet.col(1).width = 4500
    sheet.write(0, 2, 'Дата отправки письма', style_title)
    sheet.col(2).width = 7000
    sheet.write(0, 3, 'Заголовок письма', style_title)
    sheet.col(3).width = 10000
    sheet.write(0, 4, 'Текст письма', style_title)
    sheet.col(4).width = 15000

    for i, order in list(enumerate(order_list, 1)):

        sheet.write(i, 0, order['a_number'])
        sheet.write(i, 1, OrderState(order['slm_state']).title)
        mail_list = order['mail_list']
        if len(mail_list) > 0:
            last_mail = mail_list[-1]
            sheet.write(i, 2, from_gpdatetime(last_mail['a_date_sent']), style_date)
            sheet.write(i, 3, last_mail['a_subject'])
            sheet.write(i, 4, last_mail['a_body_part'].strip())

    filename = 'orders_{}.xls'.format(timestamp.strftime('%y%m%d-%H'))
    path = os.path.join('data', filename)
    wb.save(path)

    # print('Отправка отчета...')
    # msg = MIMEMultipart('mixed')
    # msg['Subject'] = 'Актуальные заявки {}'.format(timestamp.strftime(timestamp_format))
    # msg['From'] = addr_from
    # msg['To'] = ', '.join(addr_to)
    # msg['cc'] = ', '.join(addr_cc)
    #
    # with open(path, 'rb') as fp:
    #     part = MIMEBase('application', 'vnd.ms-excel')
    #     part.set_payload(fp.read())
    #     encoders.encode_base64(part)
    # part.add_header('Content-Disposition', 'attachment', filename=filename)
    # msg.attach(part)
    #
    # with SMTP(host=_host, port=_port) as smtp:
    #     smtp.starttls()
    #     smtp.login(_user, _pass)
    #     smtp.sendmail(addr_from, addr_to + addr_cc, msg.as_string())
    #
    # print('Отчет отправлен.')


def test_objects():

    order = Order(100, 'W11111', OrderState(0))
    # order = Order()

    print([field.title for field in order.__dict__.values()])

    # order.ref.val = 200
    # order.ref.val = 300
    # print(order.ref.val)
    # print(order.ref.title)
    # print(order.ref)
    # order.number.val = 'W22222'
    # order.state.val = OrderState(1)
    # print(order.number)
    #
    # order2 = Order(500, 'W22222', OrderState(1))
    # # order2 = Order(500)
    # order2.number.val = 'W1111'
    # order2.state = OrderState(0)
    #
    # print(order.number.val)
    # print(order.ref.val)
    #
    # print(order2.number.val)
    # print(order2.ref.val)

def sent_mail():
    print('Отправка письма...')
    msg = MIMEText('Тестовое письмо')
    msg['Subject'] = 'Тестовое письмо'
    msg['From'] = 'sd-exch1 <sd-exch1@sberbank.ru>'
    msg['To'] = 'bespalov@diasoft-service.ru'
    # msg['cc'] = ', '.join(addr_cc)

    # with open(path, 'rb') as fp:
    #     part = MIMEBase('application', 'vnd.ms-excel')
    #     part.set_payload(fp.read())
    #     encoders.encode_base64(part)
    # part.add_header('Content-Disposition', 'attachment', filename=filename)
    # msg.attach(part)

    with SMTP(host=_host, port=_port) as smtp:
        smtp.starttls()
        smtp.login(_user, _pass)
        smtp.sendmail(addr_from, addr_to + addr_cc, msg.as_string())

    print('Письмо отправлено.')
