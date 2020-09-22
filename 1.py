import pandas as pd
import xlwings as xw
import numpy as np

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)
""" выделяем таблицу без шапки. 0-компания, 1-отступы, 2-аналитика,
3,4,5,6,7,8-суммы от начДт до конКт"""
xlapp = xw.apps.active
rng = xlapp.selection
dp = pd.DataFrame(rng.raw_value)
dp = dp.fillna(0)
dp.to_csv('KD.csv', index=False)

# dp = pd.read_csv('KD1.csv')
#
# print(dp)
# print(dp.dtypes)