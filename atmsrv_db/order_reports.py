from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP

import xlwt as xlwt

from atmsrv_db.gptyp import from_gpdatetime, OrderState
from atmsrv_db.orcl import Orcl

sqltext_mail = """
select om.a_order_number, m.a_mail_ref, m.a_date_ins, m.a_date_sent, m.a_subject, m.a_recipients, mb.a_part_num, mb.a_body_part  from  r_send_mail m
  inner join r_order_mail om on m.a_mail_ref = om.a_mail_ref
  inner join r_mail_body_part mb on mb.a_mail_ref = m.a_mail_ref
where 1 = 1
  and om.a_order_number = :order_number
  and regexp_like(m.a_recipients, 'ncrmailbox')
"""

sqltext_order_list ="""
select * from r_order o
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


def actual_ncr():
    db = Orcl()
    # db = Orcl(uri="prom_ust_atm/121@fast")

    print('Получение списка актуальных заявок...')
    order_list = db.exec_in_dict(sqltext_order_list, {})

    # print(order_list)

    for order in order_list:
        mail_list = db.exec_in_dict(sqltext_mail, {'order_number': order['a_number']})
        order['mail_list'] = sorted(mail_list, key=lambda send: send['a_date_sent'])


    # for order in order_list:
    #     print(order)

    # print(order_list)

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

    filename = 'data/orders_{}.xls'.format(timestamp.strftime('%y%m%d-%H'))
    wb.save(filename)

    print('Отправка отчета...')
    msg = MIMEMultipart('mixed')
    msg['Subject'] = 'Актуальные заявки {}'.format(timestamp.strftime(timestamp_format))
    msg['From'] = addr_from
    msg['To'] = ', '.join(addr_to)
    msg['cc'] = ', '.join(addr_cc)

    with open(filename, 'rb') as fp:
        part = MIMEBase('application', 'xml')
        part.set_payload(fp.read())
        encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(part)

    with SMTP(host=_host, port=_port) as smtp:
        smtp.starttls()
        smtp.login(_user, _pass)
        smtp.sendmail(addr_from, addr_to + addr_cc, msg.as_string())

    print('Отчет отправлен.')