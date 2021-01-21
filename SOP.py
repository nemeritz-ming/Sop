#  pip install openpyxl
#  need 3 files: 本地SKU和SPU对应表.xls, 在线SKU和SPU对应表.xls, SPU品类系列对应表.xls
import pandas as pd
import datetime as dt
import numpy as np
import requests
import datetime
import os
import xlsxwriter
import xlrd

# 设置时间
start_days = '2020-08-01'
end_days = '2020-10-01'
start_day = datetime.datetime.strptime(start_days, '%Y-%m-%d').date()
end_day = datetime.datetime.strptime(end_days, '%Y-%m-%d').date()
daytime = -1
if start_day <= end_day:
    daytime0 = end_day - start_day
    daytime = int(daytime0.days) + 1
else:
    print('起始日期大于结束日期')
    quit()

# 请求订单数据
data_dd = None
url = 'https://erp.banmaerp.com/Order/Order/ExportOrderHandler'

# 以天为单位来取，再合并
data_dd_by_day_list = []
for single_date in (start_day + datetime.timedelta(n) for n in range(daytime)):
    headers = {
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
        'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
    data = 'filter=%7B%22OriginalOrderTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.0000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.9999%22%2C%22Sort%22%3A-1%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22Addresses%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%7D&details%5B0%5D%5BFieldID%5D=37&details%5B0%5D%5BSort%5D=1&details%5B0%5D%5BFieldExportName%5D=&details%5B1%5D%5BFieldID%5D=38&details%5B1%5D%5BSort%5D=2&details%5B1%5D%5BFieldExportName%5D=&details%5B2%5D%5BFieldID%5D=39&details%5B2%5D%5BSort%5D=3&details%5B2%5D%5BFieldExportName%5D=&details%5B3%5D%5BFieldID%5D=40&details%5B3%5D%5BSort%5D=4&details%5B3%5D%5BFieldExportName%5D=&details%5B4%5D%5BFieldID%5D=41&details%5B4%5D%5BSort%5D=5&details%5B4%5D%5BFieldExportName%5D=&details%5B5%5D%5BFieldID%5D=44&details%5B5%5D%5BSort%5D=6&details%5B5%5D%5BFieldExportName%5D=&details%5B6%5D%5BFieldID%5D=46&details%5B6%5D%5BSort%5D=7&details%5B6%5D%5BFieldExportName%5D=&details%5B7%5D%5BFieldID%5D=47&details%5B7%5D%5BSort%5D=8&details%5B7%5D%5BFieldExportName%5D=&details%5B8%5D%5BFieldID%5D=48&details%5B8%5D%5BSort%5D=9&details%5B8%5D%5BFieldExportName%5D=&details%5B9%5D%5BFieldID%5D=49&details%5B9%5D%5BSort%5D=10&details%5B9%5D%5BFieldExportName%5D=&details%5B10%5D%5BFieldID%5D=50&details%5B10%5D%5BSort%5D=11&details%5B10%5D%5BFieldExportName%5D=&details%5B11%5D%5BFieldID%5D=51&details%5B11%5D%5BSort%5D=12&details%5B11%5D%5BFieldExportName%5D=&details%5B12%5D%5BFieldID%5D=53&details%5B12%5D%5BSort%5D=13&details%5B12%5D%5BFieldExportName%5D=&details%5B13%5D%5BFieldID%5D=217&details%5B13%5D%5BSort%5D=14&details%5B13%5D%5BFieldExportName%5D=&details%5B14%5D%5BFieldID%5D=62&details%5B14%5D%5BSort%5D=15&details%5B14%5D%5BFieldExportName%5D=&details%5B15%5D%5BFieldID%5D=65&details%5B15%5D%5BSort%5D=16&details%5B15%5D%5BFieldExportName%5D=&details%5B16%5D%5BFieldID%5D=66&details%5B16%5D%5BSort%5D=17&details%5B16%5D%5BFieldExportName%5D=&details%5B17%5D%5BFieldID%5D=67&details%5B17%5D%5BSort%5D=18&details%5B17%5D%5BFieldExportName%5D=&details%5B18%5D%5BFieldID%5D=68&details%5B18%5D%5BSort%5D=19&details%5B18%5D%5BFieldExportName%5D=&details%5B19%5D%5BFieldID%5D=70&details%5B19%5D%5BSort%5D=20&details%5B19%5D%5BFieldExportName%5D=&type=1'.format(
        single_date, single_date)
    r = requests.post(url=url, headers=headers, data=data)
    file_name = '/Users/edz/Documents/{0}订单数据.xlsx'.format(single_date)
    with open(file_name, 'wb') as file:
        file.write(r.content)
    data_dd_by_day_list.append(file_name)
    if data_dd is None:
        try:
            data_dd = pd.read_excel(file_name)
        except Exception as e:
            pass
    else:
        try:
            data_dd_cur = pd.read_excel(file_name)
            data_dd = pd.concat([data_dd, data_dd_cur], ignore_index=True)
        except Exception as e:
            pass

file_name_dd = '/Users/edz/Documents/{0}到{1}订单数据.xlsx'.format(start_day, end_day)
data_dd.to_excel(file_name_dd)

# 删除多余订单数据文件
for dir_file in data_dd_by_day_list:
    os.remove(dir_file)

# 请求采购单数据
url = 'https://erp.banmaerp.com/Purchase/Sheet/ExportPurchaseHandler'
data_cgd = None
begin_day = '2020-08-01'
begin_day = datetime.datetime.strptime(begin_day, '%Y-%m-%d').date()
diff_days = end_day - begin_day
diff_days = int(diff_days.days) + 1
months = diff_days / 30

data_cgd_by_day_list = []
headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
temp_date = begin_day
if months > 0:
    step = 30
    for single_date in (begin_day + datetime.timedelta(n) for n in range(30, diff_days, step)):
        data = 'filter=%7B%22UpdateTime%22%3A%7B%22Sort%22%3A%22-1%22%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%2C%22CreateTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.998%22%7D%7D'.format(
            temp_date, single_date)
        r = requests.post(url=url, headers=headers, data=data)
        file_name = '/Users/edz/Documents/{0}到{1}采购单数据.xlsx'.format(temp_date, single_date)
        temp_date = single_date
        with open(file_name, 'wb') as file:
            file.write(r.content)
        data_cgd_by_day_list.append(file_name)
        if data_cgd is None:
            data_cgd = pd.read_excel(file_name)
            data_cgd = pd.DataFrame(data_cgd.iloc[1:].values, columns=data_cgd.iloc[0, :])
        else:
            try:
                data_cgd_cur = pd.read_excel(file_name)
                data_cgd_cur = pd.DataFrame(data_cgd_cur.iloc[1:].values, columns=data_cgd_cur.iloc[0, :])
                data_cgd = pd.concat([data_cgd, data_cgd_cur], ignore_index=True)

            except Exception as e:
                continue
data = 'filter=%7B%22UpdateTime%22%3A%7B%22Sort%22%3A%22-1%22%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%2C%22CreateTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.998%22%7D%7D'.format(
            temp_date, end_day)
r = requests.post(url=url, headers=headers, data=data)
file_name = '/Users/edz/Documents/{0}到{1}采购单数据.xlsx'.format(temp_date, end_day)
with open(file_name, 'wb') as file:
        file.write(r.content)
data_cgd_by_day_list.append(file_name)
if data_cgd is None:
    data_cgd = pd.read_excel(file_name)
    data_cgd = pd.DataFrame(data_cgd.iloc[1:].values, columns=data_cgd.iloc[0, :])
else:
    try:
        data_cgd_cur = pd.read_excel(file_name)
        data_cgd_cur = pd.DataFrame(data_cgd_cur.iloc[1:].values, columns=data_cgd_cur.iloc[0, :])
        data_cgd = pd.concat([data_cgd, data_cgd_cur], ignore_index=True)
    except Exception as e:
        pass

file_name_cgd = '/Users/edz/Documents/采购单数据.xlsx'
data_cgd.to_excel(file_name_cgd)
# 删除多余订单数据文件
for dir_file in data_cgd_by_day_list:
    os.remove(dir_file)
data_cgd = pd.read_excel(file_name_cgd)
# data_cgd = pd.read_excel('/Users/edz/Documents/2021-01-20 16_26采购单.xls')
# data_cgd = pd.DataFrame(data_cgd.iloc[1:].values, columns=data_cgd.iloc[0, :])

# 请求库存数据
url = 'https://erp.banmaerp.com/Stock/SelfInventory/ExportHandler'
headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
data = 'filter=%7B%22Quantity%22%3A%7B%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A10000%2C%22PageNumber%22%3A1%7D%7D'
r = requests.post(url=url, headers=headers, data=data)
file_name_kc = '/Users/edz/Documents/库存数据.xls'
with open(file_name_kc, 'wb') as file:
    file.write(r.content)

# 请求在线商品数据
url = 'https://erp.banmaerp.com/Shopify/Product/ExportHandler'
headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
data = 'filter=%7B%22UpdateTime%22%3A%7B%22Sort%22%3A-1%7D%7D'
r = requests.post(url=url, headers=headers, data=data)
file_name_cp = '/Users/edz/Documents/在线商品数据.xlsx'
with open(file_name_cp, 'wb') as file:
    file.write(r.content)

# sop3
# 读取库存数据，并根据是否有标题找到我们需要的行列。
data_kc = pd.read_excel(file_name_kc)
if "库存清单数据" in data_kc.columns.tolist():
    data_kc = pd.DataFrame(data_kc.iloc[1:].values, columns=data_kc.iloc[0, :])
# 只取 坑头仓库+虹猫蓝兔仓库 库存数据
data_kc = data_kc[(data_kc['仓库'] == '坑头') | (data_kc['仓库'] == '虹猫蓝兔动漫有限公司')]
# 计算出空闲库存=合格总量-合格锁定
data_kc['sku空闲库存'] = data_kc['合格总量'] - data_kc['合格锁定量']
# 读取产品数据
data_cp = pd.read_excel(file_name_cp)
if "Shopify产品" in data_cp.columns.tolist():
    data_cp = pd.DataFrame(data_cp.iloc[1:].values, columns=data_cp.iloc[0, :])

# 提取采购单sku-spu对应表
sku2spu_kc = data_cgd[['本地SKU', 'SPU']].groupby(['本地SKU'],as_index= False).first()
data_kc_combined = data_kc[['本地sku', 'sku空闲库存', '均价（￥）']].groupby(['本地sku'], as_index= False).first()

# 把sku改为spu后 再次透视得出spu空闲库存
data_kc_combined = data_kc_combined.merge(sku2spu_kc, left_on='本地sku', right_on='本地SKU', how='left')


a = data_kc_combined[['sku空闲库存', 'SPU']].groupby(['SPU'], as_index = False).sum()
a.rename(columns = {'SPU':'对应SPU'}, inplace = True)

# 读取订单数据
data_dd = pd.read_excel(file_name_dd)
# 处理下单时间，只精确到日
data_dd['下单时间'] = data_dd['下单时间'].dt.date
# 透视得出按下单日期分布的售卖件数 （独立sheet）
data_dd[['下单时间', '数量']].groupby(['下单时间']).sum()
# 分离表格成 KOL订单和 普通订单 （按照支付金额0的是KOL订单）
data_dd_kol = data_dd[(data_dd['支付金额(USD)'] == 0)]
data_dd_pp = data_dd[(data_dd['支付金额(USD)'] != 0)]
# 分别透视出 KOL订单的售卖件数（源b） 和 普通订单的售卖件数（源c） sku为行
b0 = data_dd_kol[['匹配SKU', '数量']].groupby(['匹配SKU'], as_index=False).sum()
c0 = data_dd_pp[['匹配SKU', '数量']].groupby(['匹配SKU'], as_index=False).sum()
# 提取在线sku-spu对应表
sku2spu_cp = data_cp[['SPU', 'Sku']]
# 把sku改成SPU 再透视出spu维度的统计
b0 = b0.merge(sku2spu_cp, left_on='匹配SKU', right_on='Sku', how='left')
# 提取错误sku

b = b0[['数量', 'SPU']].groupby(['SPU'], as_index=False).sum()
b.rename(columns= {'SPU':'SPU-B'}, inplace = True)


c0 = c0.merge(sku2spu_cp, left_on='匹配SKU', right_on='Sku', how='left')
c = c0[['数量', 'SPU']].groupby(['SPU'], as_index=False).sum()
c.rename(columns= {'SPU':'SPU-C'}, inplace = True)


# 删除在线商品列表SPU为空的行
data_cp = data_cp[~data_cp['SPU'].isnull()]

# 删除spu重复行
# print(data_cp.shape[0])
data_cp = data_cp.groupby('SPU').first().reset_index()
data_cp.rename(columns = {'SPU':'SPU_x'}, inplace = True)
# print(data_cp.shape[0])

# 处理 取最早的发布时间，只精确到日
data_cp['上架天数'] = dt.date.today() - data_cp['发布时间'].dt.date


# 补充完整Shopify分类列
SPU_category = pd.read_excel('/Users/edz/Documents/SPU品类系列对应表.xls')
data_cp = data_cp.merge(SPU_category, left_on='SPU_x', right_on='SPU', how='left')
# data_cp.loc[data_cp['Shopify分类'].isnull(), 'Shopify分类'] = data_cp['品类']
data_cp.loc[data_cp['Shopify分类'].isnull(), 'Shopify分类'] = data_cp['品类']
data_cp = data_cp[~data_cp['Shopify分类'].isnull()]

# 补充完整'系列'列
data_cp['系列_x'] = data_cp['系列_y']
data_cp = data_cp[~data_cp['系列_x'].isnull()]

# 取spu列，'shopify分类'列，'系列'列，'上架天数'列，'售卖状态'列，'图片链接'列，'售价'列，'成本价'列 得到源d
d = data_cp[['SPU_x', 'Shopify分类', '系列_x', '上架天数', '售卖状态', 'Sku图片', '价格','成本价']]
d = d.groupby('SPU_x').first().reset_index()
d.rename(columns = {'成本价':'商品成本价'}, inplace = True)
data_kc_combined = data_kc_combined[['SPU','均价（￥）']].groupby(['SPU'], as_index= False).first()
data_kc_combined.rename(columns = {'SPU':'SPU_kc'}, inplace = True)
d = d.merge(data_kc_combined, left_on = 'SPU_x', right_on = 'SPU_kc', how = 'left')
d['成本价'] = d['均价（￥）']/6.5
d['spu_上架'] = np.where(d['售卖状态'] == '上架', 1, 0)
# print(d.columns)

d = d.rename(columns={'价格': '售价'})
# print(d['Shopify分类'].unique())

# sop4
# 以源d为为基础根据源a v_lookup出空闲库存（ending inventory）
d = d.merge(a, left_on='SPU_x', right_on='对应SPU', how='left')
d = d.rename(columns={'sku空闲库存': '空闲库存'})

d = d.merge(b, left_on='SPU_x', right_on='SPU-B', how='left')
d = d.rename(columns={'数量': 'KOL订单售卖件数'})

d = d.merge(c, left_on='SPU_x', right_on='SPU-C', how='left')
d = d.rename(columns={'数量': '普通订单售卖件数'})

d['空闲库存'] = d['空闲库存'].fillna(0)
d['普通订单售卖件数'] = d['普通订单售卖件数'].fillna(0)
d['KOL订单售卖件数'] = d['KOL订单售卖件数'].fillna(0)
d['成本价'] = d['成本价'].fillna(d['商品成本价'])

#  设置"beginning inventory"列 =ending inventory+KOL订单售卖件数+普通订单售卖件数
#  设置"动销"列 =IF(kol售卖件数+普通订单售卖件数>0,1,0)
#  设置"售价和"列 =普通订单售卖件数*售价
#  设置"成本和"列 =普通订单售卖件数*成本价
#  设置"毛利率"列 =（售价和-成本和）/售价和
#  设置"日均销量"列=普通订单售卖件数/上架天数
#  设置"月存销比"列=空闲库存/（普通订单售卖件数+KOL售卖件数）/（分析时间天数/30）
d['beginning inventory'] = d['空闲库存'] + d['KOL订单售卖件数'] + d['普通订单售卖件数']
d['动销'] = np.where(d['KOL订单售卖件数'] + d['普通订单售卖件数'] > 0, 1, 0)
d['售价和'] = d['普通订单售卖件数'] * d['售价']
d['成本和'] = d['普通订单售卖件数'] * d['成本价']
d['毛利率'] = np.where(d['售价和'] != 0 | ~d['售价和'].isnull(),(d['售价和'] - d['成本和']).astype(float) / d['售价和'], 0 )
d['日均销量'] = d['普通订单售卖件数'].astype(float) /np.where(d['上架天数'].dt.days.astype('float') > daytime, daytime, d['上架天数'].dt.days.astype('float'))
d['月存销比'] = np.where(d['普通订单售卖件数'] + d['KOL订单售卖件数'] == 0, 0,
                     d['空闲库存'].astype(float) / (d['普通订单售卖件数'] + d['KOL订单售卖件数']).astype(float) / (daytime / 30.0))
d['售罄率'] = (d['普通订单售卖件数'] + d['KOL订单售卖件数']) / d['beginning inventory'].astype(float)

d['月存销比'] = d['月存销比'].fillna(0)
d['售罄率'] = d['售罄率'].fillna(0)

cur = data_cgd[['采购单号','状态', 'SPU','物品数量','到货物品数量']]
cur = cur[cur['状态'] == '采购中']
cur['在途库存'] = cur['物品数量'].astype(float) - cur['到货物品数量'].astype(float)
cur = cur[['SPU', '在途库存']].groupby(['SPU'], as_index=False).sum()
d = d.merge(cur, left_on='SPU_x', right_on='SPU', how='left')
d.loc[d['在途库存'].isnull(), '在途库存'] = 0
d['可售天数'] = (d['空闲库存'] + d['在途库存']).astype(float) / ((d['普通订单售卖件数'] + d['KOL订单售卖件数']).astype(float)/np.where(d['上架天数'].dt.days.astype('float') > daytime, daytime, d['上架天数'].dt.days.astype('float')))
d['可售天数'] = d['可售天数'].replace(np.inf, np.nan)
d = d.drop(labels=['SPU'],axis=1)

# sop5 group by 系列
d['系列_x'] = d['系列_x'].str.lower()
SPU_category['系列'] = SPU_category['系列'].str.lower()
S = d.groupby(['系列_x'], as_index=False).agg({'spu_上架': np.sum, '动销': np.sum, 'beginning inventory': np.sum,
                                             '空闲库存': np.sum, '普通订单售卖件数': np.sum, 'KOL订单售卖件数': np.sum, '日均销量': np.sum,
                                             '售价和': np.sum, '成本和': np.sum})

# sop6 group by Shopify类
# 更改shopify类别
d['Shopify分类'] = d['Shopify分类'].str.lower()
SPU_category['品类'] = SPU_category['品类'].str.lower()
d.loc[(d['Shopify分类'] == 'blazer/jacket') | (d['Shopify分类'] == 'ot') | (
        d['Shopify分类'] == 'outwear'), 'Shopify分类'] = 'Outwear'
d.loc[(d['Shopify分类'] == 'blouse') | (d['Shopify分类'] == 'top'), 'Shopify分类'] = 'Top'
d.loc[(d['Shopify分类'] == 't') | (d['Shopify分类'] == 't-shirt') | (d['Shopify分类'] == 'shirt') | (
        d['Shopify分类'] == 'sweatshirt'), 'Shopify分类'] = 'T'
d.loc[
    (d['Shopify分类'] == 'cardigan') | (d['Shopify分类'] == 'sweater') | (d['Shopify分类'] == 'sw'), 'Shopify分类'] = 'Sweater'
d.loc[(d['Shopify分类'] == 'denim') | (d['Shopify分类'] == 'dn') | (d['Shopify分类'] == '牛仔dn'), 'Shopify分类'] = 'Denim'
d.loc[(d['Shopify分类'] == 'matching set') | (d['Shopify分类'] == 'set') | (d['Shopify分类'] == 'st'), 'Shopify分类'] = 'Set'
d.loc[(d['Shopify分类'] == 'pants') | (d['Shopify分类'] == '裤子pa') | (d['Shopify分类'] == 'shorts'), 'Shopify分类'] = 'Pants'
d.loc[d['Shopify分类'] == 'acc', 'Shopify分类'] = 'Acc'
d.loc[d['Shopify分类'] == 'dress', 'Shopify分类'] = 'Dress'
d.loc[d['Shopify分类'] == 'skirt', 'Shopify分类'] = 'Skirt'
d.loc[d['Shopify分类'] == 'vest', 'Shopify分类'] = 'Vest'

SP = d.groupby(['Shopify分类'], as_index=False).agg({'spu_上架': np.sum, '动销': np.sum, 'beginning inventory': np.sum,
                                                   '空闲库存': np.sum, '普通订单售卖件数': np.sum, 'KOL订单售卖件数': np.sum,
                                                   '日均销量': np.sum,
                                                   '售价和': np.sum, '成本和': np.sum})
d['利润(USD)'] = d['售价和'] - d['成本和']
d = d[['SPU_x', 'Shopify分类', '系列_x', '上架天数', '售卖状态', 'Sku图片', '售价','成本价','空闲库存','KOL订单售卖件数',
       '普通订单售卖件数', 'beginning inventory', '动销', '售价和', '成本和', '毛利率', '日均销量', '月存销比', '售罄率',
       '在途库存', '可售天数']]
d.rename(columns={'spu_上架': 'SPU上架个数', 'Shopify分类': '品类', '系列_x': '系列', 'Sku图片': '图片链接', '售价': '售价(USD)', '成本价': '成本价(USD)',
                  ' ': '合格空闲(ending inventory)', '售价和': '售价和(USD)', '成本和': '销售成本和(USD)'}, inplace=True)
d = d.groupby('SPU_x').first().reset_index()




# 设置"SPU占比"列 =spu数/总和数
# 设置"动销率"列 =动销/SPU数
# 设置"均深"列 =beginning inventory/spu数
# 设置"售罄率"列=（普通订单售卖件数+KOL售卖件数）/beginning inventory
# 设置"月存销比"列=空闲库存/（普通订单售卖件数+KOL售卖件数）/分析时间天数/30
# 设置"平均售价"列 =售价和/普通订单售卖件数
# 设置"平均成本价"列 =成本和/普通订单售卖件数
# 设置"毛利率"列 =（售价和-成本和）/售价和
dataframe = [S, SP]
for X in dataframe:
    X['SPU占比'] = X['spu_上架'] / X['spu_上架'].sum()
    X['动销率'] = X['动销'] / X['spu_上架']
    X['均深'] = X['beginning inventory'] / X['spu_上架']
    X['售罄率'] = (X['普通订单售卖件数'] + X['KOL订单售卖件数']) / X['beginning inventory']
    X['月存销比'] = X['空闲库存'] / (X['普通订单售卖件数'] + X['KOL订单售卖件数']) / (daytime / 30.0)
    X['平均售价'] = X['售价和'] / X['普通订单售卖件数']
    X['平均成本价'] = X['成本和'] / X['普通订单售卖件数']
    X['毛利率'] = (X['售价和'] - X['成本和']) / X['售价和']
    X['销售件数占比'] = (X['KOL订单售卖件数'] + X['普通订单售卖件数']) / (X['普通订单售卖件数'].sum() + X['KOL订单售卖件数'].sum())
    X['销售额占比'] = X['售价和'] / X['售价和'].sum()
    X.loc["Total"] = X.sum()
S.loc["Total",'系列_x'] = np.nan
SP.loc["Total",'Shopify分类'] = np.nan
S.rename(columns={'spu_上架': 'SPU上架个数', '系列_x': '系列'}, inplace=True)
SP.rename(columns={'spu_上架': 'SPU上架个数'}, inplace=True)
S['月存销比'] = S['月存销比'].replace(np.inf, np.nan)
SP['月存销比'] = SP['月存销比'].replace(np.inf, np.nan)


# sop7
dd = data_dd.groupby('订单号').first().reset_index()
num_of_order = pd.Series([data_dd_pp['订单号'].nunique(), data_dd_kol['订单号'].nunique(), np.nan])
num_of_sold = pd.Series([data_dd_pp.数量.sum(), data_dd_kol.数量.sum(), np.nan])
upp = num_of_sold.divide(num_of_order, fill_value=np.nan)
pay_in_usd = pd.Series([dd['支付金额(USD)'].sum(), np.nan, np.nan])
payment_per_order = pay_in_usd.divide(num_of_order, fill_value=np.nan)
mean_product_price = pay_in_usd.divide(num_of_sold, fill_value=np.nan)
sum_sale = pd.Series(
    [(data_dd_pp['单价'] * data_dd_pp['数量']).sum(), (data_dd_kol['单价'] * data_dd_kol['数量']).sum(), np.nan])

# 获取总成本
data_dd = data_dd.merge(sku2spu_cp, left_on = '匹配SKU', right_on = 'Sku', how='left')
data_dd.loc[data_dd['SPU'].isnull(), 'SPU'] = data_dd['匹配SKU'].str.slice(stop=16)
cost = data_dd.merge(data_cp, left_on ='SPU', right_on = 'SPU_x', how = 'left')
cost_all = (cost['数量']*cost['成本价']).sum()
Product_Analysis1 = pd.DataFrame(
    {'{0}到{1}'.format(start_day, end_day): ['普通订单', 'KOL订单', np.nan], '天数': [daytime, daytime, np.nan],
     '订单数': num_of_order, '售卖件数': num_of_sold, 'Unit per order': upp,
     '实际支付USD': pay_in_usd, '客单价': payment_per_order, '商品均价USD': mean_product_price,
     '售价和USD': sum_sale, '成本USD': [cost_all, np.nan, np.nan]})

Product_Analysis2 = data_dd.groupby(['下单时间'], as_index=False).agg({'订单号': ["nunique"], '数量': np.sum})
Product_Analysis2['日均售卖件数'] = Product_Analysis2[('数量', 'sum')] / Product_Analysis2[('订单号', 'nunique')]
Product_Analysis2.rename(columns={'订单号': '总订单数', '数量': '售卖件数'}, inplace=True)


writer = pd.ExcelWriter('/Users/edz/Documents/{0}到{1}产品分析sop.xlsx'.format(start_day, end_day), engine='xlsxwriter')

# Write each dataframe to a different worksheet.
d.to_excel(writer, sheet_name='SPU数据表')
SP.to_excel(writer, sheet_name='shopify类别')
S.to_excel(writer, sheet_name='系列')
Product_Analysis1.to_excel(writer, sheet_name='订单种类')
Product_Analysis2.to_excel(writer, sheet_name='下单日期')


# Close the Pandas Excel writer and output the Excel file.
writer.save()


