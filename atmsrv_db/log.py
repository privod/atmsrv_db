import logging


def init():
    formatter = logging.Formatter('%(levelname)s %(asctime)s: %(filename)s[LINE:%(lineno)d]: %(message)s')
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    file_handler = logging.FileHandler('log.txt')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)
