#-*-coding:utf-8-*-
import sqlite3

import pandas as pd

conn=sqlite3.connect("../db/workincome1.db")

book=pd.read_sql_query("select * from ")