import os.path
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP

import xlwt as xlwt

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


class FieldInfo(object):
    def __init__(self, title=None):
        self.title = title


# class Field(object):
#     def __init__(self, info: FieldInfo, val=None):
#         self.val = val
#         self.info = info

class Field(object):
    def __init__(self, title, val=None):
        # self.inst = {}
        self.val = val
        self.title = title

    # def __get__(self, obj, type):
    #     return self.val
    #
    # def __set__(self, obj, val):
    #     self.val = val
    #
    # def __delete__(self, obj):
    #     pass

    # def get_title(self):
    #     return self.info.title


class _DomainMetaclass(type):

    # def field_cre(self):
    #     for field in fields:
    #         pass

    def __init__(cls, name, bases, attributes):

        for field in attributes.get('fields'):
            pass

        return super(_DomainMetaclass, cls).__new__(cls, name, bases, )


class Domain(object):
    # __metaclass__ = _DomainMetaclass
    # ref_info = FieldInfo('Индекс')
    fields = [
        ('ref', 'Индекс'),
    ]

    def field_cre(self, fields):
        for field in fields:
            self.__dict__[field[0]] = Field(field[1])

    def __init__(self):
        self.field_cre()
        # self.ref = Field(self.__class__.ref_info, ref)


class Mail(Domain):
    # date_sent = Field('Дата отправки письма')
    # subject = Field('Тема письма')
    # body = Field('Текст письма')

    def __init__(self,ref,  date_sent, subject, body):
        super().__init__(ref)
        self.date_sent = date_sent
        self.subject = subject
        self.body = body


class Order(Domain):
    # number = Field('Номер заявки')
    # state = Field('Статус')
    # mails = Field()

    def __init__(self, ref, number, state):
        super().__init__(ref)
        self.number = number
        self.state = state



def actual_ncr():
    db = Orcl()
    # db = Orcl(uri="prom_ust_atm/121@fast")

    print('Получение списка актуальных заявок...')
    # order_list = db.exec_in_dict(sqltext_order_list, {})
    #
    # for order in order_list:
    #     mail_list = db.exec_in_dict(sqltext_mail, {'order_number': order['a_number']})
    #     order['mail_list'] = sorted(mail_list, key=lambda send: send['a_date_sent'])

    order = Order(100, 'W11111', OrderState(0))

    o = OrderState(3)
    o.title

    db.sql_exec(sqltext_order_list, {})
    # order_list = db.fetchall()
    orders = [Order(ref, number, OrderState(state_int)) for ref, number, state_int in db.fetchall()]

    for order in orders:
        db.sql_exec(sqltext_mail, {'ref': order.ref})
        mails = [Mail(ref, date_sent, subject, body) for ref, date_sent, subject, body in db.fetchall()]
        # if len(mails) > 0:
        order.mails = sorted(mails, key=lambda mail: mail.date_sent )

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

    state = OrderState(3)
    print(state.title)

    order = Order(100, 'W11111', OrderState(0))
    # order = Order(100)
    order.ref.val = 200
    order.ref.val = 300
    print(order.ref.val)
    print(order.ref.info.title)
    print(order.ref)
    order.number = 'W22222'
    order.state = OrderState(1)
    print(order.number)

    order2 = Order(500, 'W22222', OrderState(1))
    # order2 = Order(500)
    order2.number = 'W1111'
    order2.state = OrderState(0)

    print(order.number)
    print(order.ref)

    print(order2.number)
    print(order2.ref)

