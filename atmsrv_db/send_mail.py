from datetime import datetime

from atmsrv_db.orcl import Orcl

sqltext = """
select om.A_ORDER_NUMBER, m.a_mail_ref, m.A_DATE_INS, m.A_DATE_SENT, m.A_SUBJECT, m.A_RECIPIENTS, mb.A_PART_NUM, mb.A_BODY_PART  from  r_send_mail m
  inner join R_ORDER_MAIL om on m.A_MAIL_REF = om.A_MAIL_REF
  inner join R_MAIL_BODY_PART mb on mb.a_mail_ref = m.a_mail_ref
where 1 = 1
  and om.A_ORDER_NUMBER = :order_number
  and REGEXP_LIKE(m.A_RECIPIENTS, 'ncrmailbox')
"""


def ncr_last_by_order(number_list):
    # db = Orcl()
    db = Orcl(uri="prom_ust_atm/121@fast")


    for number in number_list:
        rpt = db.exec_in_dict(sqltext, {'order_number': number})

        # for item in rpt:
        #     print(item['a_mail_ref'],  datetime.strptime(str(item['a_date_sent']),'%Y%m%d%H%M%S'), item['a_subject'], item['a_body_part'])
        # print()

        rpt = sorted(rpt, key=lambda send: send['a_date_sent'])

        if len(rpt) == 0:
            print()
            continue

        rpt_last = rpt[-1]

        rpt_text = "\t".join([
            datetime.strptime(str(rpt_last['a_date_sent']), '%Y%m%d%H%M%S').strftime('%d.%m.%Y %H:%M:%S'),
            rpt_last['a_subject'],
            '"{}"'.format(rpt_last['a_body_part'].strip()),
            rpt_last['a_order_number'],

        ])

        print(rpt_text)
