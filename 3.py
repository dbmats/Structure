import pandas as pd
import xlwings as xw

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 20)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

dp = pd.read_csv('KD1.csv')
df = dp.loc[dp['1'] == 2]
dp = dp[['2', '20']]

dd = df.merge(dp, left_on=['15'], right_on=['20'])
dd = dd.merge(dp, left_on=['16'], right_on=['20'])
# dd = dd.merge(dp, left_on=['17'], right_on=['20'])
# dd.to_csv('KD3.csv', index=False)

# dd = pd.read_csv('KD3.csv')
# dd['sch'] = (dd['2_y'].astype('str')).str[0:2]
# dd = dd[['0', '2_y.1', '2_x', '2_x.1', '2_y', '2', '3', '4', '5', '6', '7', '8']]

xlapp = xw.apps.active
rng = xlapp.selection
rng.options(index=False).value = dd
print(dd)