import pandas as pd
import xlwings as xw
import numpy as np

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

dp = pd.read_csv('KD.csv')
""" сквозные счетчики, если значение увеличивается то текущие 
счетчики следующих отступов обнуляются"""
x00 = x1 = x2 = x3 = x4 = 0
# текущии счетчики
x11 = x12 = x13 = x14 =0

for index, row in dp.iterrows():
    if row[0] > x00:
        x00 =x00 + 1
        x11 = x12 = x13 = 0
    if row[1] == 0:
        x1 = x1 + 1
        x11 = x11 + 1
        x12 = x13 = 0
        dp.loc[index, 9] = x11
        dp.loc[index, 20] = '{};{}'.format(x00, x11)
    else:
        dp.loc[index, 9] = x11

        if row[1] == 1:
            x2 = x2 + 1
            x12 = x12 + 1
            x13 = 0
            dp.loc[index, 10] = x12
            dp.loc[index, 15] = '{};{}'.format(x00, x11)
            dp.loc[index, 20] = '{};{};{}'.format(x00, x11, x12)
        else:
            dp.loc[index, 10] = x12

            if row[1] == 2:
                x3 = x3 + 1
                x13 = x13 + 1
                dp.loc[index, 11] = x13
                dp.loc[index, 15] = '{};{}'.format(x00, x11)
                dp.loc[index, 16] = '{};{};{}'.format(x00, x11, x12)
                dp.loc[index, 20] = '{};{};{};{}'.format(x00, x11, x12, x13)
            else:
                dp.loc[index, 11] = x13

                if row[1] == 3:
                    x4 = x4 + 1
                    x14 = x14 + 1
                    x15 = x16 = 0
                    dp.loc[index, 12] = x14
                    dp.loc[index, 15] = '{};{}'.format(x00, x11)
                    dp.loc[index, 16] = '{};{};{}'.format(x00, x11, x12)
                    dp.loc[index, 17] = '{};{};{};{}'.format(x00, x11, x12, x13)
                    dp.loc[index, 20] = '{};{};{};{};{}'.format(x00, x11, x12, x13, x14)
                else:
                    dp.loc[index, 12] = x14

dp = dp.fillna(0)

# xlapp = xw.apps.active
# rng = xlapp.selection
# rng.options(index=False).value = dp

print(dp)

dp.to_csv('KD1.csv', index=False)