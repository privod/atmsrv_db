from datetime import datetime
from rels import Column, Relation

_gpdatetime_format = "%Y%m%d%H%M%S"
_gpdate_format = "%Y%m%d"

def to_gpdatetime(val):
    return val.strftime(_gpdatetime_format)


def to_gpdate(val):
    return val.strftime(_gpdate_format)


def from_gpdatetime(val):
    return datetime.strptime(str(val), _gpdatetime_format)


def from_gpdate(val):
    return datetime.strptime(str(val), _gpdate_format)


class Enum(Relation):
    name = Column(primary=True)
    value = Column(external=True)
    title = Column()


class OrderState(Enum):
    records = (
        ('not_in_list', 0, 'Нет в перечне'),
        ('received', 1, 'В работе у коорд.'),
        ('appointed', 2, 'Назначена'),
        ('canceled', 3, 'Отклонена'),
        ('accepted', 4, 'Принята'),
        ('refusal', 5, 'Отказ'),
        ('wait_of_part', 6, 'Ожидание ЗИП'),
        ('order_of_part', 7, 'Заказ ЗИП'),
        ('set_off', 8, 'Выезд'),
        ('wait_of_access', 9, 'Ожидание доступа'),
        ('start_of_work', 10, 'Начало работ'),
        ('finish_of_work', 11, 'Окончание работ'),
        ('closed', 12, 'Закрыта'),
        ('repair', 13, 'В ремонте'),
        ('received_part', 14, 'Получен ЗИП'),
        ('atm_sms_1', 15, 'СМС(выключение питания)'),
        ('atm_sms_2', 16, 'СМС(включение питания)'),
        ('atm_sms_3', 17, 'СМС(питание восстановлено)'),
        ('wait_of_service', 18, 'Ожидание обслуживания'),
        ('to_once_only', 19, 'Перевод в платную'),
        ('duplicate', 20, 'Дубликат'),
        ('change_service', 21, 'Изменение типа обслуживания'),
        ('without_set_off', 22, 'Решено без выезда'),
    )

class ServiceType(Enum):
    records = (
        ('flm', 0, 'FLM'),
        ('slm_guarantee', 1, "гарантийное"),
        ('slm_postguarantee', 2, 'постгарантийное'),
        ('once_only', 3, "платное")
    )


def from_slm_state(val):
    return