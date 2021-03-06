#!/usr/bin/python
# -*- coding: utf-8 -*-
import logging

import os
os.environ['NLS_LANG'] = 'American_America.AL32UTF8'

import cx_Oracle


class Orcl:
    def __init__(self, user='privod_ust_atm_new', password='121', dns='orcl'):
        try:
            self.conn = cx_Oracle.connect(user, password, dns)
        except cx_Oracle.DatabaseError as info:
            print("DB logon  error:", info)
            exit(0)
        self.cursor = self.conn.cursor()

        def __del__(self):
            self.conn.close()

    def sql_exec(self, sql, params):
        try:
            if params:
                self.cursor.execute(sql, params)
            else:
                self.cursor.execute(sql)
        except cx_Oracle.DatabaseError as info:
            logging.error('SQL error: {}\n\tSQL execute: {}\n\tParams: {}'.format(info, self.cursor.statement, params))

    def fetchall(self):
        return self.cursor.fetchall()

    def fetchone(self):
        return self.cursor.fetchone()

    def get_metadata(self):
        return self.cursor.description

    def get_result(self):
        # desc_full = [desc + (title,) for desc, title in zip(self.cursor.description, titles)]
        metadata = [item for item in zip(*self.cursor.description)]
        return Result(metadata[0], metadata[1], self.cursor.fetchall())

    def exec_in_dict(self, sql, params):
        self.sql_exec(sql, params)

        titles = [desc[0].lower() for desc in self.cursor.description]
        dicts = []
        for rows in self.cursor.fetchall():
            dict = {}
            for title, value in zip(titles, rows):
                dict[title] = value
            dicts.append(dict)
        return dicts

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

class Result:
    def __init__(self, rows, names, types, titles=None):
        self._rows = rows
        self._names = names
        self._types = types
        if titles is None:
          self._titles = names
        else:
          self._titles = titles

    def set_titles(self, titles):
        self._titles = titles