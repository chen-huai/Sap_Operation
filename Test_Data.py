import pandas as pd
import numpy as np
import os
os.chdir('C:\\Users\\chen-fr\\Desktop\\临时文件\\sap')
#读入数据
rawData = pd.read_excel('20230331.xlsx')
#定义拼接函数，并对字段进行去重
def concat_func(x):
    return pd.Series({
        '合并':'\n'.join(x['合并'].unique()),
        # '性别':','.join(x['性别'].unique())
    }
    )
#分组聚合+拼接
# result=rawData.groupby(rawData["Invoices' name (Chinese)"]).apply(concat_func).reset_index()
# result = rawData.groupby(["Invoices' name (Chinese)",'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month']).sum(['Amount', 'Amount with VAT', 'Total Cost', 'Revenue\n(RMB)'])

rawData['合并']  = rawData['Project No.'] + '\t' + rawData['Currency']

combineProject = rawData.groupby(["Invoices' name (Chinese)",'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month']).apply(concat_func).reset_index()

result = pd.merge(rawData, combineProject, on=["Invoices' name (Chinese)",'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month'],how='right')

# result['Test'] =  'Payment Notice\nPO' + '\n' + result['合并']

#结果展示
result.to_csv('result.csv', encoding='utf_8_sig')
