#!/usr/bin/python
# -*- coding: utf-8 -*-

import cx_Oracle

class Orcl:
  def __init__(self, uri = "privod_ust_atm_new/121@orcl"):
    try:
      self.conn = cx_Oracle.connect(uri)
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
      print("SQL execute:", self.cursor.statement)
      print("Params:", params)
      print("SQL error:", info)
      input(" - = ! = - ")
       
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
