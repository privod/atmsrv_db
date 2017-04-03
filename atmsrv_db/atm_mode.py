import csv
import logging
import os.path

import re
from datetime import datetime

from atmsrv_db.gptyp import to_gpdatetime

import atm
import atmsrv_db.log

from atmsrv_db.orcl import Orcl

sqltext_order_list = """
select ref, a_number, atm_ref, atm_city, atm_addr, atm_location from r_order o
where 1 = 1
    and o.service_type = 0
    and o.date_reg = (select max(ol.date_reg) from r_order ol where ol.atm_ref = o.atm_ref)
"""

sqltext_mail = """
select m.num, m.chunk from r_mail_chunk m
where m.order_ref = :o_ref
"""

sqltext_atm = """
select * from r_atm_arc aa
    inner join r_atm a on aa.a_atm_arc = a.a_atm_arc
where aa.ref = :atm_ref
"""


def get_mode_by_order(mode_day):
    atm_mode_re = re.search('(\d\d):(\d\d)\s*-\s*(\d\d):(\d\d)', mode_day)
    if atm_mode_re:
        return int(atm_mode_re.group(1)) * 100 + int(atm_mode_re.group(2)), int(atm_mode_re.group(3)) * 100 + int(atm_mode_re.group(4))


def get_atm_last_order():
    db = Orcl()
    # db = Orcl(user='prom_ust_atm', password='121', dns='fast')

    mails = []

    logging.info('Получение списка последних заявок по каждому УС...')

    # d = timestamp.date()
    # dt_beg = to_gpdatetime(datetime.combine(d, time()))
    # dt_end = to_gpdatetime(datetime.combine(d, time(23, 59, 59)))

    db.sql_exec(sqltext_order_list, {})

    for ref, number, atm_ref, atm_city, atm_addr, atm_location in db.fetchall():
        # print(number)
        db.sql_exec(sqltext_mail, {'o_ref': ref})

        chunks = [chunk for m_num, chunk in db.fetchall() if chunk]

        mail = ''.join(chunks)

        service_win_re = re.search('Сервисное окно: ([\d:\-,\s]+)', mail)
        if service_win_re:
            service_win = service_win_re.group(1)
            service_win = re.sub('\s', '', service_win)
            logging.info(': '.join([number, service_win]))

            # db.sql_exec(sqltext_atm, {'atm_ref': atm_ref})
            # serial, city, addr, location = db.fetchone()
            # print(city, addr, location)
            # print(atm_city, atm_addr, atm_location)

            service_days = re.split(',', service_win)
            if len(service_days) == 7:
                mode_mon_beg, mode_mon_end = get_mode_by_order(service_days[0])
                mode_tue_beg, mode_tue_end = get_mode_by_order(service_days[1])
                mode_wed_beg, mode_wed_end = get_mode_by_order(service_days[2])
                mode_thu_beg, mode_thu_end = get_mode_by_order(service_days[3])
                mode_fri_beg, mode_fri_end = get_mode_by_order(service_days[4])
                mode_sat_beg, mode_sat_end = get_mode_by_order(service_days[5])
                mode_sun_beg, mode_sun_end = get_mode_by_order(service_days[6])

                print (number, mode_mon_beg, mode_mon_end, mode_sun_beg, mode_sun_end)


def get_mode_by_xls(mode_day):
    atm_mode_re = re.search('(\d+)\s*-\s*(\d+)', mode_day)
    if atm_mode_re:
        return int(atm_mode_re.group(1)), int(atm_mode_re.group(2))


def set_atm_mode_csv(fname, description):
    db = Orcl()
    # db = Orcl(user='prom_ust_atm', password='121', dns='fast')

    with open(os.path.join('data', fname)) as f:
        atm_modes = [[row[0], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[1]] for row in csv.reader(f, delimiter='\t')]

    for atm_mode in atm_modes[1:]:

        atm_dict = atm.get_atm_by_number(db, atm_mode[0])
        if not atm_dict: continue

        atm_dict['mon_beg'], atm_dict['mon_end'] = get_mode_by_xls(atm_mode[1])
        atm_dict['tue_beg'], atm_dict['tue_end'] = get_mode_by_xls(atm_mode[2])
        atm_dict['wed_beg'], atm_dict['wed_end'] = get_mode_by_xls(atm_mode[3])
        atm_dict['thu_beg'], atm_dict['thu_end'] = get_mode_by_xls(atm_mode[4])
        atm_dict['fri_beg'], atm_dict['fri_end'] = get_mode_by_xls(atm_mode[5])
        atm_dict['sat_beg'], atm_dict['sat_end'] = get_mode_by_xls(atm_mode[6])
        atm_dict['sun_beg'], atm_dict['sun_end'] = get_mode_by_xls(atm_mode[7])

        atm_dict['a_arc_desc'] = description
        atm.atm_upd(db, atm_dict)
        logging.info(atm_mode)


def test():

    try:
        atmsrv_db.log.init()

        # get_atm_last_order()

        fname = 'atm_mode_20170323.csv'
        description = 'Загрузка информации о сервисном оконе [12c5ef92-eb53-a1a4-8be5-4625a8279d62@diasoft-service.ru]'
        set_atm_mode_csv(fname, description)
        logging.info('Серивисные окна загружены')

    except Exception as e:
        logging.exception(e)


