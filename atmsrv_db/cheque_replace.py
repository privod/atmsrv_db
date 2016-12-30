from atmsrv_db.gptyp import to_gpdatetime
from atmsrv_db.orcl import Orcl


descr_list = ['ЗАКРЫТ (РАСХ.МАТ)', 'МАЛО РАСХ. МАТ', 'СБОЙ ЧЕК.ПРИНТЕРА']


def col_descr(list):
    when = []
    for item in list:
        when.append("when instr(upper(descr), '{0}') <> 0 then '{0}'".format(item))

    if len(when) > 0:
        return "case {0} end std_descr".format(" ".join(when))
    else:
        return None


def get_sqltext(producer_list=[], city_list=[]):
    return """
select date_end, std_descr, count(*) cnt from (
    select trunc(o.date_end/100000000) date_end,
        {col_descr}
    from r_order o
        inner join R_CHEQUE_REPLACE c on o.ref = c.a_order
        inner join r_atm_arc aa on o.atm_ref = aa.ref
        inner join r_atm a on aa.a_atm_arc = a.a_atm_arc
    where 1 = 1
        and date_end between :dt_beg and :dt_end
        and instr(o.a_number, 'ID') <> 0
        {where_producer}
        {where_city}
)
where 1 = 1
group by date_end, std_descr
order by date_end, std_descr
""".format(
        col_descr=col_descr(descr_list),
        where_producer=where_producer(producer_list),
        where_city=where_city(city_list),
    )


def oracle_str_array(list):
    list = ["'{}'".format(item) for item in list]
    return "({})".format(", ".join(list))

def where_producer(producer_list):
    if len(producer_list) > 0:
        return "and aa.producer_name in {}".format(oracle_str_array(producer_list))


def where_city(city_list):
    if len(city_list) > 0:
        return "and aa.city in {}".format(oracle_str_array(city_list))

# db = Orcl()
db = Orcl(uri="prom_ust_atm/121@fast")


def get_next_month(date):
    try:
        return date.replace(date.year, date.month + 1, 1)
    except ValueError:
        return date.replace(date.year + 1, 1, 1)


def find(rpt, date_end, std_descr):

    for rpt_row in rpt:
        if rpt_row['date_end'] == date_end and rpt_row['std_descr'] == std_descr:
            return  rpt_row['cnt']

    return 0


def count_by_month(moth_beg, moth_end, producer_list=[], city_list=[]):

    rpt = db.exec_in_dict(get_sqltext(producer_list, city_list), {
        'dt_beg': to_gpdatetime(moth_beg),
        'dt_end': to_gpdatetime(moth_end)
    })

    # for rpt_row in rpt:
    #     print(rpt_row)

    moth = moth_beg
    while (moth < moth_end):
        gpmonth = int(moth.strftime("%Y%m"))

        for descr in descr_list:
            cnt = find(rpt, gpmonth, descr)
            print(gpmonth, descr, cnt, sep='\t')

        moth = get_next_month(moth)
