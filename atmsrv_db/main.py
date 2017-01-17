from datetime import datetime

from atmsrv_db.send_mail import ncr_last_by_order
from atmsrv_db.order_reports import actual_ncr
from atmsrv_db.gptyp import OrderState


actual_ncr()

# number_list = ['W612135154', 'W612135180', 'W612145900', 'W612156128', 'W612205226', 'W612205262', 'W612216274', 'W612226173', 'W612235274', 'W612235699', 'W612235791', 'W612236127', 'W612265335', 'W612267007', 'W612275583', 'W612276014', 'W612276027', 'W612276093', 'W612276096', 'W612276098', 'W612276275', 'W612285187', 'W612295210', 'W612295268', 'W612296334', 'W612296344', 'W612305077', 'W612305257', 'W612315077', 'W612315724', 'W701055095', 'W701056245', 'W701065118', 'W701065213', 'W701065667', 'W701075138', 'W701075141', 'W701085235', 'W701085638', 'W701095217', 'W701095300', 'W701095719', 'W701095725', 'W701095749', 'W701095755', 'W701096562', 'W701096570', 'W701096615', 'W701105142', 'W701105255', 'W701105297', 'W701106298', 'W701106661', 'W701106809', 'W701115107', 'W701115155', 'W701115309', 'W701115311', 'W701115345', 'W701115438', 'W701115580', 'W701115600', 'W701115609', 'W701115634', 'W701115654', 'W701116083', 'W701116085', 'W701116165', 'W701116397', 'W701116456', 'W701116517', 'W701116521', 'W701116804', 'W701116818', 'W701117057', 'W701117127', 'W701117145', 'W701125183', 'W701125261', 'W701125263', 'W701125430', 'W701125443', 'W701125566', 'W701125605', 'W701125736', 'W701126042', 'W701126044']
#
# ncr_last_by_order(number_list)


# TODO Move to tests
# from atmsrv_db.cheque_replace import count_by_month, get_sqltext
#
# moth_beg = datetime(2016, 1, 1)
# moth_end = datetime(2016, 12, 30, 23, 59, 59)
#
# count_by_month(
#     moth_beg, moth_end,
#     producer_list=['DSV'],
#     city_list=['Талнах', 'Норильск', 'Кайеркан'],
# )
# print()
# count_by_month(
#     moth_beg, moth_end,
#     producer_list=['NCR'],
#     city_list=['Талнах', 'Норильск', 'Кайеркан'],
# )
# print()
# count_by_month(
#     moth_beg, moth_end,
#     producer_list=['Nautilus'],
#     city_list=['Талнах', 'Норильск', 'Кайеркан'],
# )




# rb = xlrd.open_workbook("2.xlsx")
# sheet = rb.sheet_by_index(0)
# row =  sheet.row_values(2)
# c = sheet.cell(2, 1)
# print(sheet.row_types(2))
# print(row)
# print(type(row[1]))

# xf = sheet.book.xf_list[c.xf_index]
# print(xf)
# fmt_obj = sheet.book.format_map[xf.format_key]
# print(fmt_obj)
#
# print(repr(c.value), c.ctype, fmt_obj.type, fmt_obj.format_key, fmt_obj.format_str)

# print(str(int(row[1])))
# print(row[16])

# cl = ['DI', 'DO', 'AI', 'AO', 'M', 'Event', 'VALVE', 'MODULE']
# for i in cl:
#     if i in rb.sheet_names():
#         sheet = rb.sheet_by_name(i)
#         for rownum in range(1, sheet.nrows):
#             row = sheet.row_values(rownum)
#             eval(i)(row, rownum)