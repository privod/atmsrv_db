from datetime import datetime


def to_gpdatetime(val):
    return val.strftime("%Y%m%d%H%M%S")


def to_gpdate(val):
    return val.strftime("%Y%m%d")
