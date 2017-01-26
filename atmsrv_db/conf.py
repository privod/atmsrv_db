import json
import os


class Conf(object):
    _file_name = 'atmsrv_db.cfg'

    @classmethod
    def _get_file_name_bad(cls):
        return cls._file_name + '.bad'

    def __init__(self):
        file_name = self.__class__._file_name
        file_name_bad = self.__class__._get_file_name_bad()

        # try:
        s = open(file_name).read()
        self._conf = json.loads(s)
        # self.__dict__.update(self._conf)
        # except FileNotFoundError as e:
        #     raise e
        # except ValueError:
        #     os.rename(file_name, file_name_bad)

    def get(self, key):
        return self._conf.get(key)
