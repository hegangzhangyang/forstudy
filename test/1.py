# -*-coding:utf-8-*-
import sqlite3

import pandas as pd
from core.wage import Wage
# 这两段代码设置输出时对齐
import numpy as np
from lib.exceImport import ExcelToSql

pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
# 设置最大显示列数为10
pd.set_option('display.max_columns', 100)
# 设置控制台现实不允许换行
pd.set_option('expand_frame_repr', False)
conn = sqlite3.connect(r"..\db\workincome1.db")

# book = pd.read_excel("11月奖金自己手工帐.xlsx")
# book.fillna(0,inplace=True)
# book.to_sql("bouns201911", conn,index=False,if_exists="replace")


