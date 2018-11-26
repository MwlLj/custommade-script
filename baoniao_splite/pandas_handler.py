import numpy as np
import pandas as pd
# a = np.arange(24).reshape(4, 6)
# a[2][4] = 100
# a[1][4] = 100
# b = a[:, 4]
# c = a[:, 5]
# # a = a[a[:, 4] > a[:, 5]]
# a = a[b > c]
# print(a)
excel_ori = pd.read_excel(io = 'data.xlsx')
a = excel_ori.values
a = a[a[:, 7] > a[:, 8]]
data_df = pd.DataFrame(a)
 
data_df.columns = ['单据号','商品编码','商品售价','销售数量','消费金额','消费产生的时间','收银机号','实际收费','消费金额']
# data_df.index = ['a','b','c','d','e','f','g','h']
 
writer = pd.ExcelWriter('ret.xlsx')
data_df.to_excel(writer, 'page_1', index=False)
writer.save()
