import logging
from datetime import datetime

from atmsrv_db.gptyp import to_gpdatetime

sqltext_count_get = """
select a_countval from r_count c
where c.a_countname = :cnt_name
"""

sqltext_count_upd = """
update r_count c set a_countval = :cnt_val
where c.a_countname = :cnt_name
"""

sqltext_atm_sel = """
select * from r_atm a
    inner join r_atm_arc aa on a.a_atm_arc = aa.a_atm_arc
where 1 = 1
"""


def count_inc(db, cnt_name):
    db.sql_exec(sqltext_count_get, {'cnt_name': cnt_name})
    cnt_val = db.fetchone()[0]
    cnt_val += 1

    db.sql_exec(sqltext_count_upd, {'cnt_name': cnt_name, 'cnt_val': cnt_val})
    db.commit()
    return cnt_val


def atm_upd(db, atm):
    atm_arc = count_inc(db, 'r_atm_arc')
    atm['a_atm_arc'] = atm_arc
    atm['a_arc_dt'] = to_gpdatetime(datetime.now())
    atm['a_arc_usr'] = 'script'
    atm['a_arc_stat'] = 2

    # atm_values = ['\'' + str(val or '') + '\'' for val in atm.values()]
    # print(atm_values)
    cols = ', '.join(atm.keys())
    pars = ', '.join([':' + col for col in atm.keys()])
    sqltext = 'insert into r_atm_arc ({}) values({})'.format(cols, pars)
    db.sql_exec(sqltext, atm)

    sqltext = 'update r_atm set a_atm_arc = :arc where ref = :ref'
    db.sql_exec(sqltext, {'arc': atm_arc, 'ref':atm['ref']})

    db.commit()

    # print(atm_db)

def get_atm_by_number(db, num):
    sqltext = sqltext_atm_sel + 'and a_number = :num'
    atms = db.exec_in_dict(sqltext, {'num': num})

    if len(atms) == 0 :
        logging.warning('Устройсто ЦФС={} в справочнике не найдено'.format(num))
        return None
    elif len(atms) > 1 :
        logging.warning('По ЦФС={} в справочнике найдено несколько УС'.format(num))
        for atm in atms:
            logging.warning(str(atm))
        return None

    return atms[0]


