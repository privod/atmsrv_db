from datetime import datetime

from atmsrv_db.cheque_replace import count_by_month, get_sqltext




# TODO Move to tests
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